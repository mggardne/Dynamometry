%#######################################################################
%
%                      * Fatigue Torque Program *
%
%          M-File which reads a dynamometer CSV file and finds the
%     maximum (peak) fatigue torque values, time tension integrals
%     (TTIs) and maximum rates of torque development (RTDs) for each
%     cycle of the fatigue testing.
%
%          Maximum torque and rate of torque development (RTD) from a
%     prior isometric maximum voluntary contraction (MVC) trial are
%     used as potential values to normalize the fatigue contractions.
%
%          For each of the fatigue cycles, the maximum torques (Nm),
%     the rank of the maximum torques, relative torques (to the maximum
%     isometric MVC or maximum torque within the first seven cycles),
%     trial relative torques (to the maximum torque within the trial),
%     time tension integrals (TTIs) (Nms), relative TTIs (to the maximum
%     TTI within the first seven cycles), rates of torque development
%     (RTD) (Nm/s), relative RTDs (to the maximum isometric MVC or
%     maximum RTD within the first ten cycles), trial relative RTDs
%     (to the maximum RTD within the first ten cycles) are output to a
%     sheet (FatigueTorque) in the MS-Excel spreadsheet,
%     "FatigueTorque.xlsx", in the path of the data file.
%
%          Plots of the data with the maximum torques are written to a
%     PDF file with the CSV file name and PDF file extension in the path
%     of the data file.
%
%     NOTES:  1.  The output MS-Excel spreadsheet,
%             "FatigueTorque.xlsx" can NOT be open in another 
%             program (e.g. MS-Excel, text editor, etc.) while using
%             this program.
%
%             2.  The sides (left or right) of the legs of the subjects
%             should be in the CSV file names.  Left legs are coded as
%             zeros (0) and right legs are coded as ones (1) in the
%             output spreadsheet.
%
%             3.  M-files lsect3.m and lsect4.m must be in the current
%             path or directory.
%
%             4.  This program only reads the CSV files from the CRC
%             dynamometer.  The Stafford dynamometer CSV files have a
%             different format.
%
%             5.  The program uses Acrobat Distiller to convert the
%             postscipt file (*.ps) to a PDF file (*.pdf).   Acrobat
%             Distiller is assumed to be on the path:
%           "C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\acrodist"
%
%     20-Sep-2018 * Mack Gardner-Morse
%

%#######################################################################
%
% Clear Workspace
%
clc;
clear all;
close all;
fclose all;
%
% Threshold for Finding the Cycles of Maximum Torques
%
% thres = 1.5-sqrt(5)/2;  % ~38.2% of maximum
thres = 0.15;           % 15% of maximum
%
% Time for Start Check
%
strt_chk = 0.25;        % First quarter second
%
% Size of Moving Average Bin
%
bin = 9;                % Bin size for moving average (Need a 0.1 s at 100 Hz sampling to get baseline)
hbin = fix(bin/2);      % Offset for centering moving average
%
% IIR Low Pass Filter Parameters
%
cutoff = 10;            % 10 Hz cutoff
filterOrder = 2;        % (Desired filter order)/2 to account for filtfilt processing => 4th order filter
%
% Acrobat Distiller Path and Arguments
%
pdist = '"C:\Program Files (x86)\Adobe\Acrobat 11.0\Acrobat\acrodist"';
parg = ' --deletelog:on /N /Q ';       % Arguments for Acrobat Distiller
%
% Output MS_Excel Spreadsheet File Name
%
xlsnam = 'FatigueTorque.xlsx';
shtnam = 'FatigueTorque';
hdr = {'File','Date','Leg','Side','Cycle','Max Torque','Rank', ...
       'Relative Torque','Trial Relative Torque','TTI', ...
       'Relative TTI','RTD','Relative RTD','Trial Relative RTD'};    % Column headers
units = {'','','','','','(Nm)','','','','(Nms)','','(Nm/s)','',''};  % Units
%
% Matlab Version
%
mver = verLessThan('Matlab','9');      % Include -fillpage in print command on newer Matlab versions
%
% Get Subject's Isometric Maximum Voluntary Contraction (MVC) Data
%
nam = 'Input Isometric Trial Data';    % Name for input dialog box
dhdr = [['Please enter the values from the best isometric maximum', ...
         ' voluntary']; ['                    contraction trial ', ...
         'prior to the fatigue trial.']];   % Dialog header
prmpt1 = 'Maximum torque (Nm): ';      % First prompt
prmpt2 = 'Maximum rate of torque development (Nm/s): ';    % Second prompt
%
iso_data = cell(0);     % Cell array for character answers to prompts
%
while isempty(iso_data);
     iso_data = inputdlg({dhdr;prmpt1;prmpt2},nam,[0;1;1]);
     if isempty(iso_data)
       return;          % User canceled input -> exit program
     end
     iso_mxtorq = str2num(iso_data{2});     % Isometric maximum torque
     iso_rtdmx = str2num(iso_data{3});      % Isometric maximum RTD
     if isempty(iso_rtdmx)||isempty(iso_mxtorq)
       iso_data = cell(0);             % No user inputs
     end
end
%
% Get Subject's Fatigue File Name and Path
%
[fnamf,pnamf,fidx] = uigetfile({'*.csv', 'All CSV files'; ...
              '*.*', 'All files'},['Please Select Fatigue Protocol', ...
              ' Data CSV File for Analysis']);
%
if fidx==0              % User hit "Cancel" button
  return;
end
%
fnamfc = {fnamf};       % Make single file name a cell array
%
% Parse File Name for Left or Right Leg
%
sidef = 0;
leg = strfind(lower(fnamf),'left');
if isempty(leg)
  sidef = 1;
  leg = strfind(lower(fnamf),'right');
  if isempty(leg)
    while isempty(leg)
         leg = questdlg(['Did the subject use the left or right' ...
                         ' leg?'],'Please Choose a Side','Left', ...
                         'Right','Left');   % Prompt the user for leg
    end
    if strcmp(leg,'Left')
      sidef = 0;
    else
      sidef = 1;
    end
  end
end
%
% Read Data from Fatigue CRC Dynamometer CSV File
%
fid = fopen(fullfile(pnamf,fnamf),'r');
%
frmt = '"%f" "%f" "%f" "%f" "%f"';
data = textscan(fid,frmt,'Delimiter',',','HeaderLines',1); % Use csvread?
%
fclose(fid);            % Close CSV file
%
% Get Variables from the Data
%
t = data{2};            % Time (s)
t = t-t(1);             % Time starts at zero
npts = size(t,1);
torq = data{4};         % Torque (Nm)
%
% Get Sampling Rate from the Time Vector
%
samplingRate = 1./mean(diff(t));
samplingRate = round(1e+6*samplingRate)./1e+6;   % Remove any truncation errors
%
% Get Filter Coefficients
%
nyquist = samplingRate/2; 
Wn = cutoff/nyquist;
[b,a] = butter(filterOrder,Wn,'low');  % Lowpass Butterworth filter
%
% Filter Torque Data and Get Rate of Torque Development (RTD)
%
filt_torq = filtfilt(b,a,torq);
dfilt_torq = diff(filt_torq);
dfilt_torq(npts) = 0;
RTDfilt = dfilt_torq*samplingRate;     % RTD (Nm/s)
%
% Get Maximum Torque and Threshold Torque
%
[mxtorq,idmx] = max(torq);
thres_torq = thres*mxtorq;
%
% Plot Torque Data and Threshold
%
hf2 = figure('Name','Torque','Units','normalized','Position', ...
             [0 0.1 1 0.80]);
orient landscape;
%
hl = plot(t,torq,'b.-','LineWidth',1.0,'MarkerSize',7);
hold on;
%
xlabel ('Time (s)','FontSize',12,'FontWeight','bold');
ylabel('Torque (Nm)','Color','b','FontSize',12,'FontWeight','bold');
ht = title(fnamf,'FontSize',18,'FontWeight','bold','Interpreter', ...
           'none');
%
axis auto;
axlim = axis;
%
htl = plot(axlim(1:2),[thres_torq thres_torq],'k-', ...
           'LineWidth',1.5);
%
% Check If Start of Torque Data is Below the Threshold?
%
idx = 1:strt_chk*samplingRate;         % Index to start of torque data
%
if all(torq(idx)>thres_torq)
  kstrt = menu({'  Initial torque values are above the threshold!'; ...
                ['How should the start of the first cycle be ', ...
                'defined?']}, 'Use the first torque data point.', ...
                'Pick a time for the start of the initial cycle.', ...
                'Pick a new threshold.');
%
  if kstrt==1
    idc1 = 1;
    h1 = plot([t(idc1) t(idc1)],axlim(3:4),'g-','LineWidth',1.5);
    figure(hf2);
    drawnow;
%
  elseif kstrt==2
    kpck = true;
    while kpck
         uiwait(msgbox({'Pick a Start Time for the Initial Cycle.'}, ...
                        'non-modal'));
         figure(hf2);
         [itime,~] = ginput(1);
%
         idc1 = round(samplingRate*itime)+1;     % From time to index
%
% Plot Initial Time
%
         h1 = plot([t(idc1) t(idc1)],axlim(3:4),'g-','LineWidth',1.5);
         figure(hf2);
         drawnow;
%
% Check Initial Time
%
         kpck = logical(menu({'Start time OK?'; ...
                        '(vertical green line)'},'Yes','No')-1);
%
    end
%
  else
    kthr = true;
    while kthr
         uiwait(msgbox({'Pick a New Threshold Torque.'},'non-modal'));
         figure(hf2);
         [~,thres_torq] = ginput(1);
%
% Plot New Threshold Torque
%
         set(htl,'YData',[thres_torq thres_torq]);
         figure(hf2);
         drawnow;
%
% Check Threshold Torque
%
         kthr = logical(menu({'Threshold torque OK?'; ...
                        '(horizontal black line)'},'Yes','No')-1);
%
    end
  end

end
%
% Find Fatigue Loading Cycles by Looking for Intersections with a Line
% at the Threshold Torque
%
s = warning;
warning('off');
[~,~,idc] = lsect4([t(1) thres_torq],[t(end) thres_torq],[t torq]);
warning(s);
%
% Get Number of Cycles
%
if exist('idc1','var')
  idc = [idc1; idc];
  delete(h1);
end
%
ncyc = size(idc,1);
if mod(ncyc,2)          % Split into even number of beginnings and endings
  ncyc = (ncyc-1)/2;
  idc = reshape(idc(1:end-1),2,ncyc)';
else
  ncyc = ncyc/2;
  idc = reshape(idc,2,ncyc)';
end
%
% Plot Cycles for Verification
%
figure(hf2);
%
hclb = plot(repmat(t(idc(:,1))',2,1),repmat(axlim(3:4)',1,ncyc), ...
            'g-','Color',[0 0.5 0],'LineWidth',0.5);  % Cycle begin
hcle = plot(repmat(t(idc(:,2))',2,1),repmat(axlim(3:4)',1,ncyc), ...
            'r-','LineWidth',0.5);     % Cycle end
%
xpc = t(idc');
xpc = [xpc; flipud(xpc)];
ypc = reshape(repmat(axlim(3:4),2,ncyc),4,ncyc);
zpc = -ones(4,ncyc);
hpc = patch(xpc,ypc,zpc,zpc,'FaceColor',[0.9 0.95 1],'EdgeColor', ...
            'none','FaceAlpha',0.5);
%
% Add Menu to Set the Visibility of the Cycles On and Off
%
hm1 = uimenu(hf2,'Label','Cycles');
hm11 = uimenu(hm1,'Label','Visible?','Checked','on');
cstr1 = ['chk = get(hm11,''Checked''); if strcmp(chk,''on'');', ...
         'if double(hpc)>0;', ...
         'set([hclb; hcle; hpc],''Visible'',''off''); else;', ...
         'set([hclb; hcle],''Visible'',''off''); end;', ...
         'set(hm11,''Checked'',''off'');', ...
         'else; if double(hpc)>0;', ...
         'set([hclb; hcle; hpc],''Visible'',''on''); else;', ...
         'set([hclb; hcle],''Visible'',''on''); end;', ...
         'set(hm11,''Checked'',''on''); end;'];
%
set(hm11,'CallBack',cstr1);
%
% Confirm Threshold Torque
%
kthr = logical(menu({'Threshold torque OK?'; ...
               '(horizontal black line)'},'Yes','No')-1);
%
% Let User Check and Pick New Threshold for Finding Cycles
%
while kthr
%
     delete([hclb; hcle; hpc]);        % Delete cycle markers
%
     uiwait(msgbox({'Pick a New Threshold Torque.'},'non-modal'));
     figure(hf2);
     [~,thres_torq] = ginput(1);
%
% Plot New Threshold Torque
%
     set(htl,'YData',[thres_torq thres_torq]);
     figure(hf2);
     drawnow;
%
% Check If Start of Torque Data is Below the Threshold?
%
     if all(torq(idx)>thres_torq)
       kstrt = menu({['  Initial torque values are above the ', ...
                     'threshold!'];['How should the start of the ', ...
                     'first cycle be defined?']}, ...
                     'Use the first torque data point.', ...
                     ['Pick a time for the start of the initial ', ...
                     'cycle.'],'Delete start time for initial cycle.');
%
       if kstrt==1
%
         if exist('h1','var')
           delete(h1);
         end
%
         idc1 = 1;
         h1 = plot([t(idc1) t(idc1)],axlim(3:4),'g-','LineWidth',1.5);
         figure(hf2);
         drawnow;
%
       elseif kstrt==2
         kpck = true;
         while kpck
              uiwait(msgbox(['Pick a Start Time for the Initial ', ...
                     Cycle.'],'non-modal'));
              figure(hf2);
%
              if exist('h1','var')
                delete(h1);
              end
%
              [itime,~] = ginput(1);
%
              idc1 = round(samplingRate*itime)+1;     % From time to index
%
% Plot Initial Time
%
              h1 = plot([t(idc1) t(idc1)],axlim(3:4),'g-', ...
                        'LineWidth',1.5);
              figure(hf2);
              drawnow;
%
% Check Initial Time
%
         kpck = logical(menu({'Start time OK?'; ...
                        '(vertical green line)'},'Yes','No')-1);
%
         end
%
       else
         clear idc1;
       end
%
     end                % End of check of start
%
% Find Fatigue Loading Cycles by Looking for Intersections with a Line
% at the New Threshold Torque
%
     s = warning;
     warning('off');
     [~,~,idc] = lsect4([t(1) thres_torq],[t(end) thres_torq],[t torq]);
     warning(s);
%
     if exist('idc1','var')
       idc = [idc1; idc];
       delete(h1);
     end
%
     ncyc = size(idc,1);
     if mod(ncyc,2)
       ncyc = (ncyc-1)/2;
       idc = reshape(idc(1:end-1),2,ncyc)';
     else
       ncyc = ncyc/2;
       idc = reshape(idc,2,ncyc)';
     end
%
% Plot New Cycles
%
     hclb = plot(repmat(t(idc(:,1))',2,1), ...
                 repmat(axlim(3:4)',1,ncyc),'g-','Color',[0 0.5 0], ...
                 'LineWidth',0.5);     % Cycle begin
     hcle = plot(repmat(t(idc(:,2))',2,1), ...
                 repmat(axlim(3:4)',1,ncyc),'r-','LineWidth',0.5);   % Cycle end
%
     xpc = t(idc');
     xpc = [xpc; flipud(xpc)];
     ypc = reshape(repmat(axlim(3:4),2,ncyc),4,ncyc);
     zpc = -ones(4,ncyc);
     hpc = patch(xpc,ypc,zpc,zpc,'FaceColor',[0.9 0.95 1], ...
                 'EdgeColor','none','FaceAlpha',0.5);
     set(hm11,'Checked','on');
%
% Check Threshold Torque
%
     kthr = logical(menu({'Threshold torque OK?'; ...
                    '(horizontal black line)'},'Yes','No')-1);
%
end
%
% Check Cycles
%
set([hclb; hcle; hpc],'Visible','on');
set(hm11,'Checked','on');
%
kcyc = logical(menu('Fatigue cycles OK?','Yes','No')-1);
%
while kcyc
%
% Check Cycle Beginnings
%
     delete(hpc);       % Delete incorrect cycles
     set(hcle,'Color',[0.5 0.5 0.5]);  % Set inactive ending lines to gray
     figure(hf2);
     drawnow;
%
     kbeg = logical(menu('Beginning green lines OK?', ...
                    'Yes','No')-1);
%
     nbl = ncyc;   % Number of beginning lines
%
     while kbeg
%
          uiwait(msgbox({'Press "A" to add a new beginning line.'; ...
                         ['Press "D" to delete an existing ' , ...
                         'beginning line.']; ['Press "M" to move ', ...
                         'an existing beginning line.']; ...
                         'Press space when finished.'},'Message', ...
                         'none','replace'));
%
% Get Keyboard Input
%
          b = 0;
          while b~=32;
%
               [~,~,b] = ginput(1);    % Get keyboard input
%
% Add a Beginning Line
%
               if b==65||b==97         % Uppercase A or lower case a
                 uiwait(msgbox(['Please select the start of a ', ...
                                 'cycle for a new beginning line.'], ...
                                 'Message','none','replace'));
%
                 [tb,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                 nbl = nbl+1;
                 idc(nbl,1) = round(samplingRate*tb)+1;
%
% Plot New Beginning Line
%
                 hclb(nbl) = plot(repmat(t(idc(nbl,1)),2,1), ...
                                  axlim(3:4)','g-', ...
                                  'Color',[0 1 0],'LineWidth',0.5);
%
% New Line OK?
%
                 kadd = logical(menu('New beginning green line OK?', ...
                                     'Yes','No')-1);
%
                 while kadd
                      delete(hclb(nbl));
                      uiwait(msgbox(['Please select the start of ', ...
                                     'a cycle for a new beginning ', ...
                                     'line.'],'Message','none', ...
                                     'replace'));
%
                      [tb,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                      idc(nbl,1) = round(samplingRate*tb)'+1;
%
% Plot New Beginning Line
%
                      hclb(nbl) = plot(repmat(t(idc(nbl,1)),2,1), ...
                                       axlim(3:4)','g-', ...
                                       'Color',[0 1 0], ...
                                       'LineWidth',0.5);
%
% New Line OK?
%
                      kadd = logical(menu(['New beginning green ', ...
                                           'line OK?'],'Yes','No')-1);
%
                 end
%
                 set(hclb(nbl),'Color',[0 0.5 0]);    % Update color
                 figure(hf2);
                 drawnow;
%
               end
%
% Delete a Beginning Line
%
               if b==68||b==100;       % Uppercase D or lower case d
                 uiwait(msgbox(['Please select a beginning line ', ...
                                'to delete.'], ...
                                'Message','none','replace'));
%
                 [td,~] = ginput(1);
                 [~,idd] = min((t(idc(:,1))-td).^2);
%
% Check Line to Delete
%
                 set(hclb(idd),'Color','r');
                 figure(hf2);
                 drawnow;
                 kdel = logical(menu({'Delete beginning line?'; ...
                                      '     (red line)'}, ...
                                      'Yes','No')-1);
%
                 while kdel
                      set(hclb(idd),'Color',[0 0.5 0]);
                      figure(hf2);
                      drawnow;
                      uiwait(msgbox(['Please select a beginning ', ...
                                     'line to delete.'], ...
                                     'Message','none','replace'));
%
                      [td,~] = ginput(1);
                      [~,idd] = min((t(idc(:,1))-td).^2);
%
% Check Line to Delete
%
                      set(hclb(idd),'Color','r');
                      figure(hf2);
                      drawnow;
                      kdel = logical(menu({'Delete beginning line?'; ...
                                           '     (red line)'}, ...
                                           'Yes','No')-1);
%
                 end
%
% Delete Line and Update Indices
%
                 delete(hclb(idd));
                 figure(hf2);
                 drawnow;
%
                 idl = true(nbl,1);    % Logical index
                 idl(idd) = false;
                 idc = idc(idl,:);     % Update cycle index (both beginning and ending)
                 hclb = hclb(idl);     % Update beginning cycle handles
                 if idd<=ncyc
                   idl = true(ncyc,1); % Logical index
                   idl(idd) = false;
                   delete(hcle(idd));  % Delete corresponding ending line
                   hcle = hcle(idl);   % Update ending cycle handles
                 end
                 nbl = nbl-1;          % Reduce the number of beginnings
                 ncyc = ncyc-1;        % Reduce the number of cycles
%
               end
%
% Move a Beginning Line
%
               if b==77||b==109;       % Uppercase M or lower case m
                 uiwait(msgbox(['Please select a beginning line ', ...
                                'to move.'], ...
                                'Message','none','replace'));
                 [tm,~] = ginput(1);
                 [~,idm] = min((t(idc(:,1))-tm).^2);
%
% Check Line to Move
%
                 set(hclb(idm),'Color',[0 1 0]);
                 figure(hf2);
                 drawnow;
                 kmov = logical(2-menu({'Move this beginning line?'; ...
                                        '   (bright green line)'}, ...
                                       'Yes','No'));
%
                 while kmov
                      uiwait(msgbox(['Please select the start of ', ...
                                     'a cycle for this beginning ', ...
                                     'line to move to.'],'Message', ...
                                     'none','replace'));
%
                      [tm,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                      idc(idm,1) = round(samplingRate*tm)+1;
%
% Move Line on Plot
%
                      set(hclb(idm),'XData',repmat(t(idc(idm,1)),2,1));
                      figure(hf2);
                      drawnow;
%
% Check Move
%
                      kmov = logical(menu(['Moved beginning ', ...
                                             'line OK?'],'Yes','No')-1);
%
                 end
%
                 set(hclb(idm),'Color',[0 0.5 0]);
                 figure(hf2);
                 drawnow;
%
               end
%
          end
%
          kbeg = logical(menu('Beginning green lines OK?', ...
                              'Yes','No')-1);
%
     end
%
% Check Cycle Endings
%
     set(hcle,'Color','r');            % Make endings red
     set(hclb,'Color',[0.5 0.5 0.5]);  % Set inactive beginning lines to gray
     figure(hf2);
     drawnow;
%
% Check Number of Endings and Beginnings
%
     imsg = 'Ending red lines OK?';
     if nbl>ncyc
       kmtch = false;
       imsg = {imsg; ['NOTE:  Number of ending lines do ', ...
                      'not match the number of beginning lines.']};
     else
       kmtch = true;
     end
%
     kend = logical(menu(imsg,'Yes','No')-1);
%
     while kend||~kmtch
%
          uiwait(msgbox({'Press "A" to add a new ending line.'; ...
                         ['Press "D" to delete an existing ' , ...
                         'ending line.']; ['Press "M" to move ', ...
                         'an existing ending line.']; ...
                         'Press space when finished.'},'Message', ...
                         'none','replace'));
%
% Get Keyboard Input
%
          b = 0;
          while b~=32;
%
               [~,~,b] = ginput(1);    % Get keyboard input
%
% Add a Ending Line
%
               if b==65||b==97         % Uppercase A or lower case a
                 uiwait(msgbox(['Please select the end of a ', ...
                                 'cycle for a new ending line.'], ...
                                 'Message','none','replace'));
%
                 [te,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                 ncyc = ncyc+1;
                 idc(ncyc,2) = round(samplingRate*te)+1;
%
% Plot New Ending Line
%
                 hcle(ncyc) = plot(repmat(t(idc(ncyc,2)),2,1), ...
                                  axlim(3:4)','r-', ...
                                  'Color',[1 0.7 0.4],'LineWidth',0.5);
%
% New Line OK?
%
                 kadd = logical(menu('New ending tan line OK?', ...
                                     'Yes','No')-1);
%
                 while kadd
                      delete(hcle(ncyc));
                      uiwait(msgbox(['Please select the end of ', ...
                                     'a cycle for a new ending ', ...
                                     'line.'],'Message','none', ...
                                     'replace'));
%
                      [te,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                      idc(ncyc,2) = round(samplingRate*te)'+1;
%
% Plot New Line
%
                      hcle(ncyc) = plot(repmat(t(idc(ncyc,2)),2,1), ...
                                        axlim(3:4)','r-', ...
                                        'Color',[1 0.7 0.4], ...
                                        'LineWidth',0.5);
%
% New Line OK?
%
                      kadd = logical(menu(['New ending tan ', ...
                                           'line OK?'],'Yes','No')-1);
%
                 end
%
                 set(hcle(ncyc),'Color',[1 0 0]); % Update color
                 figure(hf2);
                 drawnow;
%
               end
%
% Delete an Ending Line
%
               if b==68||b==100;       % Uppercase D or lower case d
                 uiwait(msgbox(['Please select an ending line ', ...
                                'to delete.'], ...
                                'Message','none','replace'));
%
                 [td,~] = ginput(1);
                 [~,idd] = min((t(idc(:,2))-td).^2);
%
% Check Line to Delete
%
                 set(hcle(idd),'Color',[1 0.7 0.4]);
                 figure(hf2);
                 drawnow;
                 kdel = logical(menu({'Delete ending line?'; ...
                                      '     (tan line)'}, ...
                                      'Yes','No')-1);
%
                 while kdel
                      set(hcle(idd),'Color','r');
                      figure(hf2);
                      drawnow;
                      uiwait(msgbox(['Please select an ending ', ...
                                     'line to delete.'], ...
                                     'Message','none','replace'));
%
                      [td,~] = ginput(1);
                      [~,idd] = min((t(idc(:,2))-td).^2);
%
% Check Line to Delete
%
                      set(hcle(idd),'Color',[1 0.7 0.4]);
                      figure(hf2);
                      drawnow;
                      kdel = logical(menu({'Delete ending line?'; ...
                                           '     (tan line)'}, ...
                                           'Yes','No')-1);
%
                 end
%
% Delete Line and Update Indices
%
                 delete(hcle(idd));
%
                 n = max([nbl,ncyc]);
                 idl = true(n,1);      % Logical index
                 idl(idd) = false;
                 idc = idc(idl,:);     % Update cycle index (both beginning and ending)
                 nbl = nbl-1;          % Reduce the number of beginnings
                 ncyc = ncyc-1;        % Reduce the number of cycles
%
               end
%
% Move an Ending Line
%
               if b==77||b==109;       % Uppercase M or lower case m
                 uiwait(msgbox(['Please select an ending line ', ...
                                'to move.'], ...
                                'Message','none','replace'));
                 [tm,~] = ginput(1);
                 [~,idm] = min((t(idc(:,2))-tm).^2);
%
% Check Line to Move
%
                 set(hcle(idm),'Color',[1 0.7 0.4]);
                 figure(hf2);
                 drawnow;
                 kmov = logical(2-menu({'Move this ending line?'; ...
                                        '      (tan line)'}, ...
                                        'Yes','No'));
%
                 while kmov
                      uiwait(msgbox(['Please select the end of ', ...
                                     'a cycle for this ending ', ...
                                     'line to move to.'],'Message', ...
                                     'none','replace'));
%
                      [tm,~] = ginput(1);
%
% Go From Time to Index into Data Arrays
%
                      idc(idm,2) = round(samplingRate*tm)+1;
%
% Move Line on Plot
%
                      set(hcle(idm),'XData',repmat(t(idc(idm,2)),2,1));
                      figure(hf2);
                      drawnow;
%
% Check Move
%
                      kmov = logical(menu('Moved ending line OK?', ...
                                          'Yes','No')-1);
%
                 end
%
                 set(hcle(idm),'Color','r');
                 figure(hf2);
                 drawnow;
%
               end
%
          end
%
% Check Number of Endings and Beginnings
%
          imsg = 'Ending red lines OK?';
          if nbl~=ncyc
            kmtch = false;
            imsg = {imsg; ['NOTE:  Number of ending lines do ', ...
                    'not match the number of beginning lines.']};
          else
            kmtch = true;
          end
%
          kend = logical(menu('Ending red lines OK?','Yes','No')-1);
%
     end
%
% Sort Cycle Index and Replot Cycles
%
     idc = sort(idc);
%
     delete(hclb);
     delete(hcle);
%
     hclb = plot(repmat(t(idc(:,1))',2,1), ...
                 repmat(axlim(3:4)',1,ncyc), ...
                 'g-','Color',[0 0.5 0],'LineWidth',0.5);  % Cycle begin
     hcle = plot(repmat(t(idc(:,2))',2,1), ...
                 repmat(axlim(3:4)',1,ncyc), ...
                 'r-','LineWidth',0.5);     % Cycle end
%
     xpc = t(idc');
     xpc = [xpc; flipud(xpc)];
     ypc = reshape(repmat(axlim(3:4),2,ncyc),4,ncyc);
     zpc = -ones(4,ncyc);
     hpc = patch(xpc,ypc,zpc,zpc,'FaceColor',[0.9 0.95 1], ...
                 'EdgeColor','none','FaceAlpha',0.5);
%
     kcyc = logical(menu('Fatigue cycles OK?','Yes','No')-1);
%
end
%
% Get Maximums Within each Cycle
%
mxt = zeros(ncyc,1);
idmxf = zeros(ncyc,1);
%
for k = 1:ncyc
   idx = idc(k,1):idc(k,2);
   [mxt(k),idmxf(k)] = max(torq(idx));
   idmxf(k) = idx(idmxf(k));
end
%
[idmxs,ids] = setdiff(idmxf,idmx,'stable'); % Index to cycle maximums without maximum maximum
mxts = mxt(ids);        % Cycle maximums without maximum maximum
%
% Get Fractional Ranks of Torques From Maximum to Minimum
%
mxto = sort(mxt,'descend');            % Order torques
[~,rnk] = ismember(mxt,mxto);          % Get rank with lowest index
[~,rnko] = ismember(mxt,mxto,'legacy');% Get rank with highest index
rnk = (rnk+rnko)/2;     % Average rank
%
% Get Peak Torque for Relative Torque Values
%
pktorq = max(mxt(1:7));
if pktorq<iso_mxtorq
  pktorq = iso_mxtorq;
end
%
if pktorq<mxtorq
  uiwait(msgbox({'Maximum torque value is not within the first', ...
                 '      seven cycles or the isometric MVC.'}, ...
                 'Note','none','replace'));
end
%
% Plot Maximums
%
hmx = plot3(t(idmxs),mxts,ones(1,ncyc-1),'md','LineWidth',1);
hpk = plot3(t(idmx),mxtorq,1,'rd','LineWidth',1,'MarkerFaceColor','r');
hpk1 = plot(axlim(1:2),[mxtorq mxtorq],'r:','LineWidth',0.5);
%
% Save Plot to a PS File
%
idot = strfind(fnamf,'.');
idot = idot(end);
psnam = [fnamf(1:idot) 'ps'];          % Change file extension to PS
psnam = fullfile(pnamf,psnam);         % Add path
if mver
  print('-dpsc2',psnam);         % Matlab versions before 9
else
  print('-dpsc2','-fillpage',psnam);
end
%
% Get Moving Average and Differences of the Moving Average
% (Slope/Sampling Rate)
%
tmvavg = movmean(torq,bin,'Endpoints','discard');
dtorq = diff(tmvavg);
%
% Get Rate of Torque Development (RTD) Maximums
%
rtdmx = zeros(ncyc,1);
irtdmx = zeros(ncyc,1);
idcr = zeros(ncyc,1);
%
for k = 1:ncyc
%
   if k==1
     id1 = hbin+1;
   else
     id1 = idc(k-1,2);
   end
%
   id2 = idc(k,1);
   if id2<id1
     id2 = id1+0.1*samplingRate;
   end
   id = id1:id2;
   idm = id-hbin;       % Exclude missing start of moving average
   idr = find(dtorq(idm)<1e-6);        % Use differences in moving average to find start of cycle
   if isempty(idr)&&k==1% No flat before first cycle
     idr = 1;
   elseif isempty(idr)  % No flat before this cycle (use end of previous cycle)
     idr = id(1);
   else                 % Flat before this cycle (use last flat point before cycle)
     idr = idr(end);
     idr = id(idr);     
   end
%
   id = idr:id2+0.2*samplingRate;      % Go beyond cycle boundary by 0.2 s at 100 Hz sampling rate
   [rtdmx(k) irtdmx(k)] = max(RTDfilt(id));
   irtdmx(k) = id(irtdmx(k));
   idrc(k,:) = [idr id(end)];
%
end
%
mxrtd = max(rtdmx);     % Maximum RTD
%
% Get Peak Rate of Torque Development (RTD) for Relative RTD Values
%
mxrtd10 = max(rtdmx(1:10));            % Maximum RTD within the first ten cycles
pkrtd = mxrtd10;
if pkrtd<iso_rtdmx
  pkrtd = iso_rtdmx;
end
%
if pkrtd<mxrtd
  uiwait(msgbox({'Maximum RTD value is not within the', ...
                 ' first ten cycles or the isometric MVC.'}, ...
                 'Note','none','replace'));
end
%
% Plot Rate of Torque Development (RTD)
%
hf3 = figure('Name','RTD','Units','normalized','Position', ...
             [0 0.1 1 0.80]);
orient landscape;
%
hdl = plot(t,RTDfilt,'b.-','LineWidth',1,'MarkerSize',7);
hold on;
hml = plot(t(hbin+1:end-hbin-1),dtorq*samplingRate,'c.-', ...
           'LineWidth',0.5,'MarkerSize',7);
%
xlabel ('Time (s)','FontSize',12,'FontWeight','bold');
ylabel('RTD (Nm/s)','Color','b','FontSize',12,'FontWeight','bold');
ht = title(fnamf,'FontSize',18,'FontWeight','bold','Interpreter', ...
           'none');
%
axis auto;
axlim2 = axis;
%
hdmx = plot(t(irtdmx),rtdmx,'rd','LineWidth',1,'MarkerSize',7);
hdmxl = plot(axlim(1:2),[mxrtd mxrtd],'r:','LineWidth',0.5);
%
% hdcl0 = plot(repmat(t(idc(:))',2,1),repmat(axlim2(3:4)',1,2*ncyc), ...
%              'k-','LineWidth',1.5);
hdcl1 = plot(repmat(t(idrc(:,1))',2,1),repmat(axlim2(3:4)',1,ncyc), ...
             'g:','LineWidth',0.5);
hdcl2 = plot(repmat(t(idrc(:,2))',2,1),repmat(axlim2(3:4)',1,ncyc), ...
             'r:','LineWidth',0.5);
%
hll = legend([hdl,hml],{'Low-pass filtered RTD','Moving Average RTD'});
set(hll,'FontSize',12,'FontWeight','bold');
%
% Save Plot to a PS File
%
if mver
  print('-dpsc2','-append',psnam);    % Matlab versions before 9
else
  print('-dpsc2','-append','-fillpage',psnam);
end
%
% Calculate Time Tension Integral (TTI) for each Cycle
%
tti = zeros(ncyc,1);
%
for k = 1:ncyc
   tti(k) = sum((torq(idc(k,1):idc(k,2)-1)+ ...
                 torq(idc(k,1)+1:idc(k,2)))/2)/samplingRate;
end
%
[mxtti idt] = max(tti); % Get maximum TTI
%
% Get Peak Time Tension Integral (TTI) for Relative TTI Values
%
pktti = max(tti(1:7));
%
if pktti<mxtti
  uiwait(msgbox(['Maximum TTI value is not within the first', ...
                 ' seven cycles.'],'Note','none','replace'));
end
%
% Plot Time Tension Integrals (TTIs)
%
xp = 1:ncyc;
%
hf4 = figure('Name','TTI','Units','normalized','Position', ...
             [0 0.1 1 0.80]);
orient landscape;
%
htti = plot(xp',tti,'bo-','LineWidth',1.5,'MarkerFaceColor','b');
%
hold on;
hmxtl = plot([1; ncyc],[mxtti; mxtti],'r:','LineWidth',1);
%
bfit = polyfit((1:ncyc)',tti,1);       % Best fit line
ttil = polyval(bfit,[1;ncyc]);
htfit = plot([1; ncyc],ttil,'b:','LineWidth',1);
%
plot(idt,mxtti,'rd','LineWidth',1,'MarkerSize',8);
%
axlim3 = axis;
if axlim3(4)<1.1*mxtti
  axlim3(4) = 1.1*mxtti;
end
axis([1 ncyc 0 axlim3(4)]);
%
xlabel('Cycle Number','FontSize',12,'FontWeight','bold');
ylabel('\color{blue}Time Tension Integral (Nms)','FontSize',12, ...
       'FontWeight','bold');
ht = title(fnamf,'FontSize',18,'FontWeight','bold','Interpreter', ...
           'none');
%
% Save Plot to a PS File
%
if mver
  print('-dpsc2','-append',psnam);    % Matlab versions before 9
else
  print('-dpsc2','-append','-fillpage',psnam);
end
%
% Convert PS File to a PDF File Using Acrobat Distiller
%
eval(['!' pdist parg '"' psnam '"']);  % Command to computer OS
eval(['!del "' psnam '"']);         % Delete PS file
%
% Get Output MS-Excel Spreadsheet File Path and Name
%
fullxlsnam = fullfile(pnamf,xlsnam);
%
% Check for Output MS-Excel Spreadsheet
%
if ~exist(fullxlsnam)
  irow = 1;
  xlswrite(fullxlsnam,hdr,shtnam,['A' int2str(irow)]);
  irow = irow+1;
  xlswrite(fullxlsnam,units,shtnam,['A' int2str(irow)]);
  irow = irow+1;
else
  [~,fshtnams] = xlsfinfo(fullxlsnam);        % Get sheet names in file
  idl = strcmp(shtnam,fshtnams);    % Sheet already exists in file?
  if all(~idl)       % Sheet name not found in file
    irow = 1;
    xlswrite(fullxlsnam,hdr,shtnam,['A' int2str(irow)]);
    irow = irow+1;
    xlswrite(fullxlsnam,units,shtnam,['A' int2str(irow)]);
    irow = irow+1;
  else               % Sheet in the file
    [~,txt] = xlsread(fullxlsnam,shtnam);
    irow = size(txt,1)+1;
  end
end
%
% Write Results to MS-Excel Spreadsheet
%
leg = {'Left'; 'Right'};
c = clock;
d = {sprintf('%.0f/%.0f/%.0f',c([2 3 1]))};
tdat = [fnamfc d leg(sidef+1)];
tdat = repmat(tdat,ncyc+3,1);
rdat1 = repmat(sidef,ncyc+3,1);
rdat2 = [(1:ncyc)' mxt rnk mxt./pktorq mxt./mxtorq tti tti./mxtti ...
         rtdmx rtdmx./pkrtd rtdmx./mxrtd10];
xlswrite(fullxlsnam,tdat,shtnam,['A' int2str(irow)]);
xlswrite(fullxlsnam,rdat1,shtnam,['D' int2str(irow)]);
xlswrite(fullxlsnam,rdat2,shtnam,['E' int2str(irow)]);
%
irow = irow+ncyc;
tdat = {'Maximum'; 'Isometric MVC/Fit Slope'; 'Peak Values'};
rdat = [mxtorq NaN thres_torq NaN mxtti NaN mxrtd; ...
        iso_mxtorq NaN thres_torq./mxtorq NaN bfit(1) NaN ...
        iso_rtdmx; pktorq NaN NaN NaN pktti NaN pkrtd];
xlswrite(fullxlsnam,tdat,shtnam,['E' int2str(irow)]);
xlswrite(fullxlsnam,rdat,shtnam,['F' int2str(irow)]);
%
return
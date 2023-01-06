%#######################################################################
%
%        * Isometric Rate of Torque Development (RTD) Program *
%
%          M-File which reads dynamometer CSV files and finds the
%     isometric rate of torque development (RTD).
%
%          For the CRC dynamometer, only one maximum value is found.  
%     For the Stafford dynamometer, the isometric torques are divided
%     into cycles based on the "End Pnt 0" column to find maximum
%     values for each cycle.
%
%          The maximum torque (Nm), the maximum of the maximum torques
%     (Nm), the velocity at the maximum torque (degrees/s), position at
%     the maximum torque (degrees), maximum RTD (Nm/s), average RTD
%     (Nm/s) to 25% of peak torque, average RTD (Nm/s) to 50% of peak
%     torque and average RTD (Nm/s) between 25% and 75% of peak torque
%     are output to a sheet (IsometricRTD) in the MS-Excel spreadsheet,
%     "IsometricRTD.xlsx" in the path of the data files.
%
%          Plots of the data with the maximum torques and RTDs are
%     written to PDF files with the CSV file names and PDF file
%     extension.
%
%     NOTES:  1.  The output MS-Excel spreadsheet,"IsometricRTD.xlsx"
%             can NOT be open in another program (e.g. MS-Excel, text
%             editor, etc.) while using this program.
%
%             2.  Isometric torque CSV files should have " 55 " or
%             " 70 " in the CSV file names.
%
%             3.  The CSV file names should start with either "CRC" or
%             "Tim".
%
%             4.  The sides (left or right) of the legs of the subjects
%             should be in the CSV file names.  Left legs are coded as
%             zeros (0) and right legs are coded as ones (1) in the
%             output spreadsheet.
%
%             5.  The onset (initial or baseline) torque is defined as
%             2% of the maximum torque.
%
%     25-Oct-2017 * Mack Gardner-Morse
%
%     04-Jan-2023 * Mack Gardner-Morse - Added readmatrix command to
%     read old (2014) CSV files from the CRC dynamometer.
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
% Output MS_Excel Spreadsheet File Name
%
xlsnam = 'IsometricRTD.xlsx';
shtnam = 'IsometricRTD';
hdr = {'File','Date','Leg','# Max','Max #','Torque','Max Torque', ...
       'Velocity','Position','Peak RTD','RTD25','RTD50','RTD75-25'}; % Column headers
units = {'','','','','','(Nm)','(Nm)','(deg/s)','(deg)','(Nm/s)', ...
         '(Nm/s)','(Nm/s)','(Nm/s)'};  % Units
%
% IIR Low Pass Filter Parameters
%
cutoff = 10;            % 10 Hz cutoff
filterOrder = 2;        % (Desired filter order)/2 to account for filtfilt processing => 4th order filter
%
% Thresholds for Calculating Average RTD Values
% See Calculate Average RTDs below
%
onset_ratio = 0.02;     % Onset to peak torque ratio (2%)
peakq = 0.25;           % 25% (quarter) of peak torque
peakh = 0.50;           % 50% (half) of peak torque
peak3q = 0.75;          % 75% (3 quarters) of peak torque
threshlds = [onset_ratio; peakq; peakh; peak3q]; % Vector of torque thresholds
nt = size(threshlds,1); % Number of thresholds
%
% Y Axis Minimum
%
ymim = -50;             % Y axis minimum
%
% Matlab Version
%
mver = verLessThan('Matlab','9');      % Include -fillpage in print command on newer Matlab versions
%
% Get Input File Names
%
[fnams,pnam,fidx] = uigetfile({'* 55 *.csv;* 70 *.csv', ...
'Isometric Torque CSV files'; '*.csv', 'All CSV files'; ...
'*.*', 'All files (*.*)'},['Please Select Isometric ', ...
'Torque CSV Files for Analysis'],'MultiSelect','on');
%
if fidx==0              % User hit "Cancel" button
  return;
end
%
if iscell(fnams)
  fnams = fnams';       % Make a column vector
else
  fnams = {fnams};      % Make sure single file is a cell array
end
%
nfiles = size(fnams,1); % Number of CSV files
%
if nfiles<1
  return;
end
%
% Check File Names
%
if fidx>1               % Did not use file filter in uigetfile
%
  f55 = strfind(fnams,' 55 ');         % Check for " 55 " in file name
  f70 = strfind(fnams,' 70 ');         % Check for " 70 " in file name
%
  for k = 1:nfiles
     if ~(~isempty(f55{k})||~isempty(f70{k}))
       lidx = menu({['One or more file names do NOT contain either', ...
                     ' " 55 " or " 70 "!'];'Continue?'},'No','Yes');
       if lidx<2
         return;
       end
       break;
     end
  end
end
%
% Get Output MS-Excel Spreadsheet File Path and Name
%
fullxlsnam = fullfile(pnam,xlsnam);
%
% Check Every File for Dynamometer?
%
ichk = true;            % Check every file for type of dynamometer? (CRC or Stafford)
%
% Loop through the CSV Files
%
for k = 1:nfiles
%
% Get CSV File Name
%
   fnam = fnams{k};     % Get CSV file name
%
% Parse File Name for CRC or Stafford Dynamometer
%
   if ichk
     idyn = strncmpi('CRC',fnam,3);
     if idyn
       idyn = ~idyn;    % CRC file (not a Stafford file)
     else
       idyn = strncmpi('Tim',fnam,3);
       if ~idyn         % Neither CRC or Stafford at start of file name
         idyn = double(idyn);
         while idyn==0
              idyn = menu('Dynamometer?','CRC','Stafford');% Ask user for type of file
         end
         idyn = logical(idyn-1);
         if ~idyn
           inew = 0;
           while inew==0
                inew = menu('Old or new CSV file(s)?','Old','New');
           end
           inew = logical(inew-1);
         end
         if nfiles>1
           ichk = menu(['Are all of the files from the same ', ...
                        'dynamometer?'],'Yes','No')-1;
           ichk = logical(ichk);
         end
       end
     end
   end
%
% Parse File Name for Left or Right Leg
%
   side = 0;
   leg = strfind(lower(fnam),'left');
   if isempty(leg)
     side = 1;
     leg = strfind(lower(fnam),'right');
     if isempty(leg)
       while isempty(leg)
            leg = questdlg(['Did the subject use the left or right' ...
                             ' leg?'],'Please Choose a Side','Left', ...
                             'Right','Left');    % Prompt the user for leg
       end
       if strcmp(leg,'Left')
         side = 0;
       else
         side = 1;
       end
     end
   end
%
% Read Data from CSV File
%
   fid = fopen(fullfile(pnam,fnam),'r');
%
   if idyn
     frmt = ['"%d" %f %f %f %f %s %*s %*s %*s %*s %*s %*s ', ...
             '%*s %*s %*s %*s'];
     data = textscan(fid,frmt,'Delimiter',',','HeaderLines',1);
     data{4} = data{4}*1.35581795;     % Convert from foot-pounds to Newton-meters
   else
     if inew
       frmt = '"%f" "%f" "%f" "%f" "%f"';
       data = textscan(fid,frmt,'Delimiter',',','HeaderLines',1);
     else
       data = readmatrix(fullfile(pnam,fnam),'Delimiter',',', ...
                         'NumHeaderLines',1);    % Read "old" data files
       [nr,nc] = size(data);
       data = mat2cell(data,nr,ones(1,nc));
     end
   end
%
   fclose(fid);         % Close CSV file
%
% Get Variables from the Data
%
   t = data{2};         % Time (s)
   npts = size(t,1);
   torq = data{4};      % Torque (Nm)
   vel = data{5};       % Velocity (degrees/s)
   pos = data{3};       % Position (degrees)
%
   endpt = NaN(npts,1);
%
   if idyn
     edata = data{6};
     edata = strrep(edata,'"','');     % Remove double quotes
     idv = ~strncmp(edata,'',1);       % Find valid end point data
     endpt(idv) = str2num(char(edata));
   end
%
% Get Sampling Rate from the Time Vector
%
   samplingRate = 1./mean(diff(t));
   samplingRate = round(1e+6*samplingRate)./1e+6;          % Remove any truncation errors
%
% Get Filter Coefficients
%
   nyquist = samplingRate/2; 
   Wn = cutoff/nyquist;
   [b,a] = butter(filterOrder,Wn,'low');         % Lowpass Butterworth filter
%
% Filter Torque Data and Get Derivative of Filtered Torque Data
%
   filt_torq = filtfilt(b,a,torq);
   RTDfilt = diff(filt_torq)*samplingRate;
   tf = t(1:end-1)+1/(2*samplingRate); % Time for filtered data
%
% Get Maximum Torque
%
   [mxtorq,idmx] = max(torq);
   nmx = 1;
%
   if idyn
%
% Loop through the Cycles to Find the Maximum Torque for Each Cycle
%
     mncyc = min(endpt);
     mxcyc = max(endpt);
     mxtorq = [];
     idmx = [];
%
     for l = mncyc:mxcyc
        ide = find(endpt==l);          % Find index to this part of the cycle
        if ~isempty(ide)
          if length(ide)>2             % Avoid edge effects
            [mxtorqc,idmxc] = max(torq(ide(2:end-1)));
            mxtorq = [mxtorq; mxtorqc];
            idmx = [idmx; ide(idmxc+1)];
          end
        end
     end
%
     nmx = size(mxtorq,1);             % Number of maximum torques
%
   end                  % End of if for Stafford dynamometer
%
% Plot All the Data
%
   hf = figure('Name',fnam,'Units','normalized','Position', ...
               [0 0.1 1 0.80]);
   orient landscape;
   [ha,h1,h6] = plotyy(t,torq,t,endpt);
   set(ha(1),'YColor','k');
   set([h1;h6],'LineWidth',1.5);
   set(h1,'Color','b');
   hold on;
   h2 = plot(t,vel,'k-','LineWidth',1.5);
   h3 = plot(t,pos,'g-','LineWidth',1.5);
   h4 = plot(t,filt_torq,'b-','Color',[0.4 0.8 1],'LineWidth',1.5);
   h5 = plot(tf,RTDfilt,'k-','Color',[0.6 0.6 0.6],'LineWidth',1.5);
   xlabel ('Time (s)','FontSize',12,'FontWeight','bold');
   ylabel({'\color{blue}Torque (Nm)';
           '\color{lightBlue}Lowpass Torque (Nm)'; ...
           '\color{gray}RTD (Nm/s)'; 
           '\color{black}Velocity (^{\circ}/s)'; ...
           '\color{green}Position (^{\circ})'},'FontSize',12, ...
           'FontWeight','bold');
   ht = title(fnam,'FontSize',24,'FontWeight','bold', ...
              'Interpreter','none');
%
% Plot Resulting Maximum Torques
%
   hmx = plot(t(idmx),mxtorq,'rd','LineWidth',1.5);
%
% Add Legend and Additional Lines
%
   if idyn
     set(ha(2),'YColor',[0 0.5 0]);
     set(h6,'Color',[0 0.5 0]);
     ylabel(ha(2),'End Points','Color',[0 0.5 0],'FontSize',12, ...
            'FontWeight','bold');
     hl = legend([h1;h2;h3;h4;h5;h6;hmx],{'Torque';'Velocity'; ...
                 'Position';'Lowpass Torque';'RTD';'EndPt'; ...
                 'Max Torque'},'Location','NorthWest','Orientation', ...
                 'horizontal');
   else
     delete(ha(2));
     hl = legend([h1;h2;h3;h4;h5;hmx],{'Torque';'Velocity'; ...
                 'Position';'Lowpass Torque';'RTD';'Max Torque'}, ...
                 'Location','NorthWest','Orientation','horizontal');
   end
   set(hl,'FontSize',12,'FontWeight','bold');
%
   axis auto;
   axlim = axis;
   if min(RTDfilt)<ymim                % Keep RTDfilt from increasing the Y range too much
     idmn = find(RTDfilt>ymim);
     axmn = floor(min(RTDfilt(idmn))/10)*10;     % Round down to nearest 10
     axlim(3) = axmn;
   end
%
   axlim(4) = max([max(torq); max(filt_torq); max(pos); max(vel); ...
                   max(RTDfilt)]);     % Get maximum data value
   axlim(4) = ceil(axlim(4)/10)*10;    % Round up to nearest 10
   axlim(4) = 1.2*axlim(4);            % Make room for legend
   axlim(4) = ceil(axlim(4)/10)*10;    % Round up to nearest 10 (redundant?)
   axis(axlim);         % Set Y axis range
   set(ha(1),'YTickMode','auto');      % Get new tick marks for new Y range
   hmxl = plot(repmat(t(idmx)',2,1),repmat(axlim(3:4)',1,nmx),'r:', ...
               'LineWidth',1);                   % Lines for maximum torques
   plot(axlim(1:2),[0 0],'k:','LineWidth',1);    % Zero grid line
%
% Confirm Valid Maximum(s)
%
   kans = menu('Maximum torque(s) OK?','Yes','No')-1;
%
% Let User Pick Regions to Find the Maximums
%
   while (kans)
        delete(hmx);
        delete(hmxl);
        set([h2; h4; h5],'Visible','off');       % Increase visibility of the raw torque data
%
        uiwait(msgbox({'Pick the Start and End of Region(s)'; ...
                       '       with Maximum Torque(s).'; ...
                       '     Hit <Enter> when Finished.'},'non-modal'));
        figure(hf);
        odd = true;
        while odd
             [tr,~] = ginput;
             ngpts = size(tr,1);
             odd = logical(mod(ngpts,2));        % Needs to be pairs of points
             if odd
               uiwait(msgbox({['Please Pick Pairs of Start and ', ...
                               'End Points!']; ...
                               '                Please Try Again.'}, ...
                               'non-modal'));
             end
        end
%
% Go From Time to Index into Data Arrays
%
        nmx = ngpts/2;  % Number of regions
        idr = zeros(nmx,2);            % Index to start (first column) and end (second column) of regions
        for l = 1:ngpts
           rmod = mod(l,2);
           ii = (l+rmod)/2;            % Row index
           jj = 2-rmod;                % Column index
           del = repmat(tr(l),npts,1)-t;         % Get differences
           del = del.*del;             % Square differences
           [~,idr(ii,jj)] = min(del);
        end
%
        if ~idyn&nmx>1
          lidx = menu('More than one maximum for a CRC file?', ...
                      'Yes','No')-1;
          if lidx
            break;
          end
          for l = 1:nmx
             idc = idr(l,1):idr(l,2);
             endpt(idc) = l-1;
          end
          idyn = true;
        end
%
% Get Maximum Torque(s)
%
        mxtorq = zeros(nmx,1);
        idmx = zeros(nmx,1);
        for l = 1:nmx
           idxr = idr(l,1):idr(l,2);
           [mxtorq(l),idmx(l)] = max(torq(idxr));
           idmx(l) = idxr(idmx(l));
        end
%
% Plot Resulting Maximum Torques
%
        hmx = plot(t(idmx),mxtorq,'rd','LineWidth',1.5);
        hmxl = plot(repmat(t(idmx)',2,1),repmat(axlim(3:4)',1,nmx), ...
                    'r:','LineWidth',1);
        set([h2; h4; h5],'Visible','on');        % Restore rest of the data

%
% Check Maximum(s)
%
        kans = menu('Maximum torque(s) OK?','Yes','No')-1;
%
   end
%
% Maximum Values
%
   mxvel = vel(idmx);
   mxpos = pos(idmx);
%
% Plot Velocity(ies) and Position(s) at Maximum Torque(s)
%
   hmxv = plot(t(idmx),mxvel,'kd','LineWidth',1.5);
   hmxp = plot(t(idmx),mxpos,'gd','LineWidth',1.5);
%
% Loop through Regions (Cycles) with Maximum Torques
%
   torqt = NaN(nmx,nt); % Torque thresholds (missing values are NaNs)
   idtts = zeros(nmx,nt);              % Index to torque thresholds
   mxRTD = NaN(nmx,1);  % Maximum RTDs (missing values are NaNs)
   idrtd = zeros(nmx,1);               % Index to maximum RTDs
   tthreshlds = mxtorq*threshlds';     % Convert threshold ratios to torques
   mxendpts = endpt(idmx);             % Endpoint numbers with maximum torques
%
   for l = 1:nmx
      mxendpt = mxendpts(l);           % Endpoint index number for this cycle
      if idyn
        ide = find(endpt==mxendpt);    % Find index to all of this cycle
        ide = ide(2);   % Avoid initial edge effects
      else
        ide = 2;        % Avoid initial edge effects
      end
%
% Get Torque Thresholds
%
      idr = fliplr(ide:idmx(l));       % Index from maximum torque to initial torque for this cycle
      torqr = torq(idr);               % Maximum torque to initial torque for this cycle
      for m = 1:nt
         idtt = find(torqr<tthreshlds(l,m));     % Index to torque thresholds
         if ~isempty(idtt)
           idtt = idtt(1)-1;
           torqt(l,m) = torqr(idtt);
           idtts(l,m) = idr(idtt);
         end
      end
%
% Get Maximum RTDs
%
      [mxRTD(l),idx] = max(RTDfilt(idr));
      idrtd(l) = idr(idx);
%
   end
%
% Plot Threshold Torques
%
   trtd = NaN(nmx,nt);  % Times at torque thresholds (NaNs for missing values)
   idv = idtts~=0;      % Index to valid thresholds
   trtd(idv) = t(idtts(idv));          % Times at torque thresholds
   har1 = plot(trtd(:,1),torqt(:,1),'cd','LineWidth',1.5); % Light blue (cyan) for start of torque increase
   har2 = plot(trtd(:,2:nt),torqt(:,2:nt),'bd','LineWidth',1.5);
%
% Calculate Average RTDs
% Columns of avgRTDs are:
%   1.  RTD25 = (Torque@25% peak-Torque@2% peak)/(Time@25% peak-Time@2% peak)
%   2.  RTD50 = (Torque@50% peak-Torque@2% peak)/(Time@50% peak-Time@2% peak)
%   3.  RTDmid = (Torque@75% peak-Torque@25% peak)/(Time@75% peak-Time@25% peak)
%
   dtorq = torqt(:,2:4)-[repmat(torqt(:,1),1,2) torqt(:,2)];    % Differences in torques
   dt = trtd(:,2:4)-[repmat(trtd(:,1),1,2) trtd(:,2)];     % Differences in time
   avgRTDs = dtorq./dt; % Average RTDs
%
% Plot Maximum RTD
%
   hmxr = plot(tf(idrtd),mxRTD,'kd','Color',[0.6 0.6 0.6], ...
               'LineWidth',1.5);
   hmxrl = plot(repmat(tf(idrtd)',2,1),repmat(axlim(3:4)',1,nmx), ...
                'k:','Color',[0.6 0.6 0.6],'LineWidth',1);
%   pause;
%
% Save Plot to a PDF File
%
   idot = strfind(fnam,'.');
   idot = idot(end);
   pdfnam = [fnam(1:idot) 'pdf'];      % Change file extension to PDF
   if mver
     print('-dpdf',fullfile(pnam,pdfnam));       % Matlab versions before 9
   else
     print('-dpdf','-fillpage',fullfile(pnam,pdfnam));
   end
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
   tdat = [fnams(k) {date}];
   tdat = repmat(tdat,nmx,1);
   mxmxtorq = repmat(max(mxtorq),nmx,1);    % Maximums of maximum torques
   rdat = [repmat(side,nmx,1) repmat(nmx,nmx,1) (1:nmx)' mxtorq ...
           mxmxtorq mxvel mxpos mxRTD avgRTDs];
   xlswrite(fullxlsnam,tdat,shtnam,['A' int2str(irow)]);
   xlswrite(fullxlsnam,rdat,shtnam,['C' int2str(irow)]);
%
% Close Figure
%
   close(hf);
%
end                     % End of looping through CSV files
%
return
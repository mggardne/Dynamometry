%#######################################################################
%
%                  * Maximum Dynamic Torque Program *
%
%          M-File which reads dynamometer CSV files and finds the
%     maximum (peak) dynamic torque value(s).
%
%          For the CRC dynamometer, only one maximum value is found.  
%     For the Stafford dynamometer, the dynamic torques are divided
%     into cycles based on the "End Pnt 0" column and positive
%     velocities to find maximum values for each cycle.
%
%          The maximum torque (Nm), the maximum of the maximum torques
%     (Nm), the velocity (degrees/s) at the maximum torque, position
%     (degrees) at the maximum torque and power (W) at the maximum
%     torque are output to a sheet (DynamicTorqueCRC or 
%     DynamicTorqueStafford) in the MS-Excel spreadsheet,
%     "MaxDynamicTorque.xlsx" in the path of the data files.
%
%          Plots of the data with the maximum torques are written to
%     PDF files with the CSV file names and PDF file extension.
%
%     NOTES:  1.  The output MS-Excel spreadsheet,
%             "MaxDynamicTorque.xlsx" can NOT be open in another 
%             program (e.g. MS-Excel, text editor, etc.) while using
%             this program.
%
%             2.  Dynamic torque CSV files should have " 60 " or " 180 "
%             in the CSV file names.
%
%             3.  The CSV file names should start with either "CRC" or
%             "Tim".
%
%             4.  The sides (left or right) of the legs of the subjects
%             should be in the CSV file names.  Left legs are coded as
%             zeros (0) and right legs are coded as ones (1) in the
%             output spreadsheet.
%
%     20-Oct-2017 * Mack Gardner-Morse
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
xlsnam = 'MaxDynamicTorque.xlsx';
hdr = {'File','Date','Leg','# Max','Max #','Torque','Max Torque', ...
       'Velocity','Target Velocity','Position','Power'};   % Column headers
units = {'','','','','','(Nm)','(Nm)','(deg/s)','(deg/s)','(deg)', ...
         '(W)'};        % Units
%
% Tolerance on Target Velocity
%
vtol = 2;               % Tolerance on target velocity (degrees/s)
%
% Matlab Version
%
mver = verLessThan('Matlab','9');      % Include -fillpage in print command on newer Matlab versions
%
% Get Input File Names
%
[fnams,pnam,fidx] = uigetfile({'* 60 *.csv;* 180 *.csv', ...
'Dynamic Torque CSV files'; '*.csv', 'All CSV files'; ...
'*.*', 'All files (*.*)'},['Please Select Maximum Dynamic ', ...
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
  f60 = strfind(fnams,' 60 ');         % Check for " 60 " in file name
  f180 = strfind(fnams,' 180 ');       % Check for " 180 " in file name
%
  for k = 1:nfiles
     if ~(~isempty(f60{k})||~isempty(f180{k}))
       lidx = menu({['One or more file names do NOT contain either', ...
                     ' " 60 " or " 180 "!'];'Continue?'},'No','Yes');
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
         if nfiles>1
           ichk = menu(['Are all of the files from the same ', ...
                        'dynamometer?'],'Yes','No')-1;
           ichk = logical(ichk);
         end
       end
     end
   end
%
% Parse File Name for Target Velocity
%
   if ~isempty(strfind(fnam,' 60 '));  % Check for " 60 " in file name
     targetv = 60;
   elseif ~isempty(strfind(fnam,' 180 '));       % Check for " 180 " in file name
     targetv = 180;
   else
     lidx = 0;
     while lidx==0
          lidx = menu({'File name does NOT contain'; ...
                       ' either " 60 " or " 180 "!'; ...
                       'Target Velocity = ?'}, ...
                       '60','180','User Input','Skip','Stop');
     end
%
     switch lidx
       case 1
         targetv = 60;
       case 2
         targetv = 180;
       case 3
         targetv = {};
         while isempty(targetv)
              targetv = inputdlg('Please input the target velocity', ...
                                 'Target Velocity',1,{'60'});
              if ~isempty(targetv)
                targetv = str2num(targetv{1});
                if targetv<=0          % Positive target velocity
                  uiwait(warndlg(['Target velocity must be ', ...
                                  'greater than zero!'],['Invalid ', ...
                                  'Target Velocity'],'modal'));
                  targetv = {};
                end
              end
         end
       case 4
         break;
       otherwise
         return;
     end
%
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
   rectify = 1-2*side;  % Change the sign of the velocity for right legs
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
     frmt = '"%f" "%f" "%f" "%f" "%f"';
     data = textscan(fid,frmt,'Delimiter',',','HeaderLines',1);
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
   vel = vel*rectify;   % Reverse sign for right legs
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
          if length(ide)>10
            [mxtorqc,idmxc] = max(torq(ide));
            mxtorq = [mxtorq; mxtorqc];
            idmx = [idmx; ide(idmxc)];
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
   [ha,h1,h4] = plotyy(t,torq,t,endpt);
   set(ha(1),'YColor','k');
   set([h1;h4],'LineWidth',1.5);
   set(h1,'Color','b');
   hold on;
   h2 = plot(t,vel,'k-','LineWidth',1.5);
   h3 = plot(t,pos,'g-','LineWidth',1.5);
   xlabel ('Time (s)','FontSize',12,'FontWeight','bold');
   ylabel({'\color{blue}Torque (Nm)';
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
     set(h4,'Color',[0 0.5 0]);
     ylabel(ha(2),'End Points','Color',[0 0.5 0],'FontSize',12, ...
            'FontWeight','bold');
     hl = legend([h1;h2;h3;h4;hmx],{'Torque','Velocity','Position', ...
                 'EndPt','Max Torque'},'Location','NorthWest', ...
                 'Orientation','horizontal');
   else
     delete(ha(2));
     hl = legend([h1;h2;h3;hmx],{'Torque','Velocity','Position', ...
                 'Max Torque'},'Location','NorthWest','Orientation', ...
                 'horizontal');
   end
   set(hl,'FontSize',12,'FontWeight','bold');
   axis auto;
   axlim = axis;
   axlim(4) = 1.2*axlim(4);            % Make room for legend
   axis(axlim);         % Make room at the top for the legend
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
        nmx = ngpts/2;  % Number of regions (peaks)
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
% Check Maximum(s) are Near Target Velocity
%
   idb = abs(abs(mxvel)-targetv)>vtol; % Index to velocities not within tolerance
   mxtorq(idb) = NaN;   % If velocity target not met, set maximum torque to blank (NaN)
%
% Plot Velocity(ies) and Position(s) at Maximum Torque(s)
%
   hmxv = plot(t(idmx),mxvel,'kd','LineWidth',1.5);
   hmxp = plot(t(idmx),mxpos,'gd','LineWidth',1.5);
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
% Calculate Power
%
   mxvelrad = mxvel*pi/180;            % Convert velocity from degrees/s to rad/s
   mxpower = mxtorq.*mxvelrad;
%
% Get MS-Excel Spreadsheet Sheet Names
%
   if idyn
     shtnam = 'DynamicTorqueStafford';
   else
     shtnam = 'DynamicTorqueCRC';
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
   mxmxtorq = NaN(nmx,1);              % Maximums of extensor and flexor maximum torques
   idp = mxvel>0;       % Extensor torques
   mxmxtorq(idp) = max(mxtorq(idp));   % Maximum of extensor maximum torques
   mxmxtorq(~idp) = max(mxtorq(~idp)); % Maximum of flexor maximum torques
   rdat = [repmat(side,nmx,1) repmat(nmx,nmx,1) (1:nmx)' mxtorq ...
           mxmxtorq mxvel repmat(targetv,nmx,1) mxpos mxpower];
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
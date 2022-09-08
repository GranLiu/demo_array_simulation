% -------------------------------------------------------------------------
% Export CST result file of compact array simulation.
% 
% Reference
% [1] CST Help Contents. "Automation and Scripting - Visual Basic"
% [2] 刘燚. space.bilibili.com/55116086/channel/seriesdetail?sid=1346616
% 
% Yongxi Liu, Xi'an Jiaotong University, 2022-08.
% -------------------------------------------------------------------------
clc;
clear;
close all;

%% Set simulation params
x_num = 1;
y_num = 3; % 16, 32
path = pwd;
mkdir('data_dir');
mkdir('data_gain');

%% Initialize CST application
cst = actxserver('CSTStudio.application');  % Load CST app
mws = invoke(cst, 'Active3D');              % Link to current working MWS project
app = invoke(mws, 'GetApplicationName');    % Get current app name
ver = invoke(mws, 'GetApplicationVersion'); % Get current version no.

%% export efficiency result
for idx = 1 : x_num*y_num
    sCommand = '';
    sCommand = [sCommand 10 'SelectTreeItem("Farfields\farfield (f=2) [' num2str(idx) ']")'];
    sCommand = [sCommand 10 'With FarfieldPlot'];
    sCommand = [sCommand 10 '.Plottype("3d")'];
    sCommand = [sCommand 10 '.SetPlotMode("directivity")'];
    sCommand = [sCommand 10 '.Step(1)'];
    sCommand = [sCommand 10 '.Plot'];
    sCommand = [sCommand 10 'End With'];
    step = ['show [' num2str(idx) '] directivity'];
    invoke(mws, 'AddToHistory',step, sCommand);
    sCommand = '';
    sCommand = [sCommand 10 'With ASCIIExport'];
    sCommand = [sCommand 10 '.Reset'];
    sCommand = [sCommand 10 '.FileName ("' path '\data_dir\' num2str(idx) '.txt")'];
    sCommand = [sCommand 10 '.Execute'];
    sCommand = [sCommand 10 'End With'];
    step = ['export [' num2str(idx) '] directivity'];
    invoke(mws, 'AddToHistory',step, sCommand);
    
    sCommand = '';
    sCommand = [sCommand 10 'SelectTreeItem("Farfields\farfield (f=2) [' num2str(idx) ']")'];
    sCommand = [sCommand 10 'With FarfieldPlot'];
    sCommand = [sCommand 10 '.Plottype("3d")'];
    sCommand = [sCommand 10 '.SetPlotMode("realized gain")'];
    sCommand = [sCommand 10 '.Step(1)'];
    sCommand = [sCommand 10 '.Plot'];
    sCommand = [sCommand 10 'End With'];
    step = ['show [' num2str(idx) '] realized_gain'];
    invoke(mws, 'AddToHistory',step, sCommand);
    sCommand = '';
    sCommand = [sCommand 10 'With ASCIIExport'];
    sCommand = [sCommand 10 '.Reset'];
    sCommand = [sCommand 10 '.FileName ("' path '\data_gain\' num2str(idx) '.txt")'];
    sCommand = [sCommand 10 '.Execute'];
    sCommand = [sCommand 10 'End With'];
    step = ['export [' num2str(idx) '] realized_gain'];
    invoke(mws, 'AddToHistory',step, sCommand);
end

%% close the project without saving export operations (avoid error when opening again)
invoke(mws,'Quit');

%% release the handle
release(cst);
release(mws);
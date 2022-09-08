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
x_num = 8;
y_num = 8; % 16, 32
path = pwd;
mkdir('data_eff');

%% Initialize CST application
cst = actxserver('CSTStudio.application');  % Load CST app
mws = invoke(cst, 'Active3D');              % Link to current working MWS project
app = invoke(mws, 'GetApplicationName');    % Get current app name
ver = invoke(mws, 'GetApplicationVersion'); % Get current version no.

%% export efficiency result
for idx = 1 : x_num*y_num
    sCommand = '';
    sCommand = [sCommand 10 'SelectTreeItem("1D Results\Efficiencies\Tot. Efficiency [' num2str(idx) ']")'];
    sCommand = [sCommand 10 'With Plot1D'];
    sCommand = [sCommand 10 '.PlotView("real")'];
    sCommand = [sCommand 10 '.Plot'];
    sCommand = [sCommand 10 'End With'];
    step = ['show [' num2str(idx) '] efficiency'];
    invoke(mws, 'AddToHistory',step, sCommand);
    sCommand = '';
    sCommand = [sCommand 10 'With ASCIIExport'];
    sCommand = [sCommand 10 '.Reset'];
    sCommand = [sCommand 10 '.FileName ("' path '\data_eff\' num2str(idx) '.txt")'];
    sCommand = [sCommand 10 '.Execute'];
    sCommand = [sCommand 10 'End With'];
    step = ['export [' num2str(idx) '] efficiency'];
    invoke(mws, 'AddToHistory',step, sCommand);
end

%% close the project without saving export operations (avoid error when opening again)
invoke(mws,'Quit');

%% release the handle
release(cst);
release(mws);
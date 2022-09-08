% -------------------------------------------------------------------------
% Generate a regular rectangular array composed of dipoles in CST for simulation.
% 
% Reference
% [1] CST Help Contents. "Automation and Scripting - Visual Basic"
% [2] 刘燚. space.bilibili.com/55116086/channel/seriesdetail?sid=1346616
% 
% Yongxi Liu, Xi'an Jiaotong University, 2022-09.
% -------------------------------------------------------------------------
clc;
clear;
close all;

%% Set array structure
n_x = 2;                                    % num of antenna in a row
n_y = 2;                                    % num of antenna in a column

%% Set simulation params
f = 2e9;                                    % operating frequency [Hz]
c = 3e8;                                    % light speed [m/s]
lambda = c/f*1e3;                           % wavelength [mm]
la = 0.465;                                 % length scaling factor
L = la*lambda;                              % total length of half-wavelength dipole [mm]
g = L/100*lambda/1e3;                       % distance for center to one adjecent end [mm]
ra = 0.005;                                 % radius scaling factor
r = ra*lambda;                              % radius of the dipole [mm]
z_in = 78.3;                                % input impedance for isolated half-wavelength dipole [ohm]
Frq = [1.9, 2.1];                           % operating frequency [GHz]

%% Initialize CST application
cst = actxserver('CSTStudio.application');  % Load CST app
mws = invoke(cst, 'NewMWS');                % Build a new MWS project
app = invoke(mws, 'GetApplicationName');    % Get current app name
ver = invoke(mws, 'GetApplicationVersion'); % Get current version no.
invoke(mws, 'FileNew');                     % Open a new CST file
path = pwd;                                 % Get current work path
filename = '\demo_array.cst';               % CST file name
fullname = [path filename];
invoke(mws, 'SaveAs', fullname, 'True');    % Save current state
invoke(mws, 'DeleteResults');               % Delete old results to avoid window in CST

%% Load params to CST
invoke(mws, 'StoreParameter','lambda',lambda);
invoke(mws, 'StoreParameter','L',L);
invoke(mws, 'StoreParameter','g',g);
invoke(mws, 'StoreParameter','r',r);

%% Initialize Units
sCommand = '';
sCommand = [sCommand 'With Units' ];
sCommand = [sCommand 10 '.Geometry "mm"']; % use 10 as newline
sCommand = [sCommand 10 '.Frequency "ghz"' ];
sCommand = [sCommand 10 '.Time "ns"'];
sCommand = [sCommand 10 'End With'] ;
invoke(mws, 'AddToHistory','define units', sCommand);

%% define monitor
sCommand = '';
sCommand = [sCommand 'With Monitor'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name "farfield (f = ' num2str(mean(Frq)) 'GHz)"'];
sCommand = [sCommand 10 '.Domain "Frequency"'];
sCommand = [sCommand 10 '.FieldType "Farfield"'];
sCommand = [sCommand 10 '.MonitorValue ' num2str(mean(Frq))];
sCommand = [sCommand 10 '.ExportFarfieldSource "False"'];
sCommand = [sCommand 10 '.Create'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory','define monitor', sCommand);

%% Set operating frequency
sCommand = '';
sCommand = [sCommand 'Solver.FrequencyRange ' num2str(Frq(1)) ',' num2str(Frq(2)) ];
invoke(mws, 'AddToHistory','define frequency range', sCommand);

%% Show bounding box
plot = invoke(mws, 'Plot');
invoke(plot, 'DrawBox', 'True');

%% Build a dipole model
% model upper cylinder of dipole
Str_Name = 'upper';
Str_Component = 'Dipole';
Str_Material = 'PEC';
sCommand = '';
sCommand = [sCommand 'With Cylinder'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name "',Str_Name, '"'];
sCommand = [sCommand 10 '.Component "', Str_Component, '"'];
sCommand = [sCommand 10 '.Material "', Str_Material, '"'];
sCommand = [sCommand 10 '.OuterRadius ', '"r"']; % to use variable in CST, we need to use "xxx"
sCommand = [sCommand 10 '.InnerRadius ', '"0.0"'];
sCommand = [sCommand 10 '.Axis ', '"y"'];
sCommand = [sCommand 10 '.Yrange ', '"g","L/2"'];
sCommand = [sCommand 10 '.Xcenter ', '"0"'];
sCommand = [sCommand 10 '.Zcenter ', '"0"'];
sCommand = [sCommand 10 '.Segments ', '"0"'];
sCommand = [sCommand 10 '.Create'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory',['define cylinder:',Str_Component,':',Str_Name], sCommand);
% model lower cylinder of dipole
Str_Name = 'lower';
Str_Component = 'Dipole';
Str_Material = 'PEC';
sCommand = '';
sCommand = [sCommand 'With Cylinder'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name "',Str_Name, '"'];
sCommand = [sCommand 10 '.Component "', Str_Component, '"'];
sCommand = [sCommand 10 '.Material "', Str_Material, '"'];
sCommand = [sCommand 10 '.OuterRadius ', '"r"']; % to use variable in CST, we need to use "xxx"
sCommand = [sCommand 10 '.InnerRadius ', '"0.0"'];
sCommand = [sCommand 10 '.Axis ', '"y"'];
sCommand = [sCommand 10 '.Yrange ', '"-L/2","-g"'];
sCommand = [sCommand 10 '.Xcenter ', '"0"'];
sCommand = [sCommand 10 '.Zcenter ', '"0"'];
sCommand = [sCommand 10 '.Segments ', '"0"'];
sCommand = [sCommand 10 '.Create'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory',['define cylinder:',Str_Component,':',Str_Name], sCommand);
% define points
sCommand = '';
sCommand = [sCommand 'Pick.PickCirclecenterFromId ','"Dipole:lower", ','"2"'];
invoke(mws, 'AddToHistory','define point of lower side', sCommand);
sCommand = '';
sCommand = [sCommand 'Pick.PickCirclecenterFromId ','"Dipole:upper", ','"1"'];
invoke(mws, 'AddToHistory','define point of upper side', sCommand);
% add discrete port
sCommand = '';
sCommand = [sCommand 'With DiscretePort'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.PortNumber ','"1"'];
sCommand = [sCommand 10 '.Type ','"SParameter"'];
sCommand = [sCommand 10 '.Label ','""'];
sCommand = [sCommand 10 '.Folder ','""'];
sCommand = [sCommand 10 '.Impedance ',num2str(z_in)];
sCommand = [sCommand 10 '.VoltagePortImpedance ', '"0.0"'];
sCommand = [sCommand 10 '.Voltage ', '"1.0"'];
sCommand = [sCommand 10 '.Current ', '"1.0"'];
sCommand = [sCommand 10 '.Monitor ', '"True"'];
sCommand = [sCommand 10 '.Radius ', '"0.0"'];
sCommand = [sCommand 10 '.SetP1 ','"True", ','"0", ',num2str(-g),', ','"0"'];
sCommand = [sCommand 10 '.SetP2 ','"True", ','"0", ',num2str(g),', ','"0"'];
sCommand = [sCommand 10 '.InvertDirection ','"False"'];
sCommand = [sCommand 10 '.LocalCoordinates ','"False"'];
sCommand = [sCommand 10 '.Wire ','""'];
sCommand = [sCommand 10 '.Position ','"end1"'];
sCommand = [sCommand 10 '.Create'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory','define port', sCommand);

%% move the dipole to left-up of the array
sCommand = '';
sCommand = [sCommand 'With Transform'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name ','"port$port1"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:lower"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:upper"'];
sCommand = [sCommand 10 '.Vector ','"-lambda/2*',num2str((n_x-1)/2),...
                            '", "lambda/2*',num2str((n_y-1)/2),'", "0"'];
sCommand = [sCommand 10 '.UsePickedPoints ','"False"'];
sCommand = [sCommand 10 '.InvertPickedPoints ','"False"'];
sCommand = [sCommand 10 '.MultipleObjects ','"False"'];
sCommand = [sCommand 10 '.GroupObjects ','"False"'];
sCommand = [sCommand 10 '.Repetitions ','"1"'];
sCommand = [sCommand 10 '.MultipleSelection ','"False"'];
sCommand = [sCommand 10 '.Transform ','"Mixed", ','"Translate"'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory','move dipole', sCommand);

%% generate array by copy the origin dipole to different place
% form the first row
sCommand = '';
sCommand = [sCommand 'With Transform'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name ','"port$port1"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:lower"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:upper"'];
sCommand = [sCommand 10 '.Vector ','"lambda/2", "0", "0"'];
sCommand = [sCommand 10 '.UsePickedPoints ','"False"'];
sCommand = [sCommand 10 '.InvertPickedPoints ','"False"'];
sCommand = [sCommand 10 '.MultipleObjects ','"True"'];
sCommand = [sCommand 10 '.GroupObjects ','"False"'];
sCommand = [sCommand 10 '.Repetitions ','"',num2str(n_x-1),'"'];
sCommand = [sCommand 10 '.MultipleSelection ','"False"'];
sCommand = [sCommand 10 '.Transform ','"Mixed", ','"Translate"'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory','copy dipole to form a row', sCommand);
% copy the first row to form the array
sCommand = '';
sCommand = [sCommand 'With Transform'];
sCommand = [sCommand 10 '.Reset'];
sCommand = [sCommand 10 '.Name ','"port$port1"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:lower"'];
sCommand = [sCommand 10 '.AddName ','"solid$Dipole:upper"'];
for i = 1:n_x-1
    sCommand = [sCommand 10 '.AddName ','"port$port',num2str(i+1),'"'];
    sCommand = [sCommand 10 '.AddName ','"solid$Dipole:lower_',num2str(i),'"'];
    sCommand = [sCommand 10 '.AddName ','"solid$Dipole:upper_',num2str(i),'"'];
end
sCommand = [sCommand 10 '.Vector ','"0", "-lambda/2", "0"'];
sCommand = [sCommand 10 '.UsePickedPoints ','"False"'];
sCommand = [sCommand 10 '.InvertPickedPoints ','"False"'];
sCommand = [sCommand 10 '.MultipleObjects ','"True"'];
sCommand = [sCommand 10 '.GroupObjects ','"False"'];
sCommand = [sCommand 10 '.Repetitions ','"',num2str(n_y-1),'"'];
sCommand = [sCommand 10 '.MultipleSelection ','"False"'];
sCommand = [sCommand 10 '.Transform ','"Mixed", ','"Translate"'];
sCommand = [sCommand 10 'End With'];
invoke(mws, 'AddToHistory','copy row to form array', sCommand);

%% define solver for each port and begin simulation one-by-one
sCommand = 'Mesh.SetCreator "High Frequency"';
invoke(mws, 'AddToHistory','define creator', sCommand);
for idx = 1 : n_x*n_y
    sCommand = '';
    sCommand = [sCommand 10 'With Solver'];
    sCommand = [sCommand 10 '.Method "Hexahedral"'];
    sCommand = [sCommand 10 '.CalculationType "TD-S"'];
    sCommand = [sCommand 10 '.StimulationPort "' num2str(idx) '"'];
    sCommand = [sCommand 10 '.StimulationMode "All"'];
    sCommand = [sCommand 10 '.SteadyStateLimit "-40"'];
    sCommand = [sCommand 10 '.MeshAdaption "False"'];
    sCommand = [sCommand 10 '.CalculateModesOnly "False"'];
    sCommand = [sCommand 10 '.SParaSymmetry "False"'];
    sCommand = [sCommand 10 '.StoreTDResultsInCache  "False"'];
    sCommand = [sCommand 10 '.FullDeembedding "False"'];
    sCommand = [sCommand 10 '.SuperimposePLWExcitation "False"'];
    sCommand = [sCommand 10 '.UseSensitivityAnalysis "False"'];
    sCommand = [sCommand 10 'End With'];
    invoke(mws, 'AddToHistory',['define solver port ' num2str(idx)], sCommand);
    % begin simulation for this port
    solver = invoke(mws, 'Solver');
    invoke(solver, 'Start');
end

%% save current operations/results and close the project
invoke(mws, 'Save');
invoke(mws, 'Quit');

%% release the handle
release(cst);
release(mws);

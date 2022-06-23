function [Inport,outport,constant] =  FindPortdemo()
path='D:\File_Pycharm\py35_prj\FindPort.slx'
open_system(path)
signal = find_system('FindPort')
Inport = find_system(signal(1),'BlockType','Inport')
outport = find_system(signal(1),'BlockType','Outport')
constant = find_system(signal(1),'BlockType','Constant')
SignalHandle = find_system(gcs,'FindAll','on','BlockType','Inport')
InportPro = get(SignalHandle(1))
InportPro.SampleTime
InportPro.Name
% Name = get_param(InportPro{1},'Name')
% Outport = find_system(signal(1),'BlockType','Outport')
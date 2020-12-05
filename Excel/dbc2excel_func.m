%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                 %           %     %        % % % %
                 %           %     %        %
                 %           % % % %        % % % %
                 %           %     %        %
                 % % %       %     %        %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% ���ܣ����DBC�ļ���ȡ����EXCEL
% ���ߣ����� 
% ʱ�䣺2020.3.10
% ���뻷����Matlab2016b
% �汾��Version 1.1
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% ����DBCҪ��:DBC�źŵ�λֻ�ܰ���a-z_A-Z_0-9/�㦸��%.*����Ϊ�޵�λ
function dbc2excel_func(Path_Name,PrjName,CarType,EditPerson,Time)
%clear all ;
% clc;
%% ��ȡDBC�ļ���Ϣ
%file = fopen('demo_my.dbc','r');
file = fopen(Path_Name,'r');
%% ��ȡDBC�ļ���Ϣ
%file = fopen('demo_my.dbc','r');
text = fscanf(file,'%c');
fclose(file);
Node_Expr = 'BU_:([\s][a-z_A-Z_0-9_]{1,}){1,}';                   %node normalizing
Node_Reg = regexp(text,Node_Expr,'tokens');               %match*.XLS file
NodeExp='([a-z_A-Z_0-9_]{1,})[\s]{0,}';
Node = regexp(Node_Reg{1},NodeExp,'tokens');   % every node 
[~,number_node]= size(Node{1}); %node number 

% statistic number of node,message and signal
Node_Sig_Expr = '(BO_[\s][-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,})([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%.*_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}){1,}';
% Node_Sig_Expr = '(BO_[\s][-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,})([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}){1,}';
Node_Sig_Reg = regexp(text,Node_Sig_Expr,'tokens');  
[~,number_message] = size(Node_Sig_Reg); %node number 
Sig_Expr = '([\s]{0,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
% Sig_Expr = '([\s]{0,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s][\"][\W\w]{0,}[\"][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
% msg send cycle time    BA_ "GenMsgCycleTime" BO_ 2566901780 1000;
Node_msg_cycle_time_expr = 'BA_[\s]["]GenMsgCycleTime["][\s]BO_[\s]([0-9]{1,})[\s]([0-9]{1,});'
Node_msg_cycle_time =  regexp(text,Node_msg_cycle_time_expr,'tokens');  
for j=1:number_message
    Message_Array{j} =  Node_Sig_Reg{1,j}{1,1};   % message's ID,NAME,node    
    Sig_Message_FullArray{j} =  Node_Sig_Reg{1,j}{1,2};  
    Sig_Array{j} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Expr,'tokens'); %Sig's Name length Bais Max Min
    [~ ,row_sig]= size(Sig_Array{1,j}); %sig number 
    row_s{j} = row_sig; 
end
sig_number=0;
for i= 1:number_message
    sig_number = sig_number+row_s{j};    %sig number
end
% classify 
% classify node 1
% All_Name
Message_Node_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
Sig_Name_Expr = '[\s]{1,}SG_[\s]([a-z_A-Z_0-9_]){1,}';
for j=1:number_message
      Message_Node{j} = regexp(Message_Array{j},Message_Node_Expr,'tokens')
end
% ��ȫ���ڵ����all��
for i=1:number_node
    All{1,i}= Node{1,1}{1,i};
end

%���౨����Ϣ
Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
Message_ID_Expr = 'BO_[\s]([0-9.]{1,})'
for j=1:number_message
    Message_Node{j} = regexp(Message_Array{j},Message_Node_Expr,'tokens') %�����ϵĽڵ�
end
All_Message_ID=[];
    for i=1:number_node
        n=1;
        for j=1:number_message
             if  strcmp(char(Node{1,1}{1,i}),char(Message_Node{j}{1}))        %�ȽϽڵ������뱨���Ͻڵ�������Ƿ���ͬ 
%                   All{i}{n,2}= regexp(Message_Array{1,j},Message_Name_Expr,'tokens');%�����ͬ��������д�����Ӧ��λ�� ���������Ͻڵ�1��������ڵ�1������ͬ����ѱ������ƹ��ൽ�ڵ�1�£�            
                     All{1,i}{n,2}= regexp(Message_Array{1,j},Message_Name_Expr,'tokens');%�����ͬ��������д�����Ӧ��λ�� ���������Ͻڵ�1��������ڵ�1������ͬ����ѱ������ƹ��ൽ�ڵ�1�£�
                     All_Message_ID{1,i}{n,2}= regexp(Message_Array{1,j},Message_ID_Expr,'tokens');%�����ͬ��������д�����Ӧ��λ�� ���������Ͻڵ�1��������ڵ�1������ͬ����ѱ������ƹ��ൽ�ڵ�1�£�             
                     n=n+1;
            end
        end
    end
    
 %����û�б��ĵĽڵ�
 all_all_all = All;
 All = [];
 all_n = 0;
[~,size_all_r] =  size(all_all_all);

 for i = 1:size_all_r
     [~,sar] = size(all_all_all{1,i})
     if(sar ~= 1)
         all_n = all_n +1;
         All{1,all_n} = all_all_all{1,i};
     end
 end
 
 
  %���Ǳ��� 
 all_all_all_mesg = All_Message_ID;
 All_Message_ID = [];
 all_m_id = 0;
 for i =1 :number_node
    [~,all_all_allmg_c] = size(all_all_all_mesg{1,i});
     if(all_all_allmg_c ~= 0)
         all_m_id = all_m_id + 1;
         All_Message_ID{1,all_m_id} = all_all_all_mesg{1,i};
     end
     
 end
 [~ ,number_node] = size(All); %���½ڵ�����
 

 
    
%�ź�������ʽ

% Sig_Expr = '([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
Sig_Name_Expr = '[\s]{1,}SG_[\s]([a-z_A-Z_0-9_]){1,}';
Sig_Start_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s]([-0-9.]){1,}[|]';
Sig_Length_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|]([-0-9.]){1,}';
Sig_Factor_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(]([-0-9.]){1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Bais_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,]([-0-9.]){1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Min_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[]([-0-9.]){1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Max_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|]([-0-9.]){1,}[\]][\s]["][a-z_A-Z_0-9/�㦸��%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
% Sig_Expr = '([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
% Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
% Sig_Name_Expr = '[\s]{1,}SG_[\s]([a-z_A-Z_0-9_]){1,}';
% Sig_Start_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s]([-0-9.]){1,}[|]';
% Sig_Length_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|]([-0-9.]){1,}';
% Sig_Factor_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(]([-0-9.]){1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
% Sig_Bais_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,]([-0-9.]){1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
% Sig_Min_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[]([-0-9.]){1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
% Sig_Max_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|]([-0-9.]){1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';

% Sig_Start_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Start_Expr,'tokens');
%�����ź���Ϣ
for j=1:number_message
    mesg_name{j} = regexp(Node_Sig_Reg{1,j},Message_Name_Expr,'tokens'); % �ҳ����б���
           for  i=1:number_node   
                 u =1;
                [d{i},r{i}] = size(All{1,i});  % �ҵ�ÿ���ڵ��±��ĵ�����
                for new=1:d{i} %���µı�������������źű������ƱȽ�
                    if  strcmp(char(All{1,i}{new,2}{1}),char(mesg_name{j}{1}{1}))        %�ȽϽڵ������뱨���Ͻڵ�������Ƿ���ͬ 
                           All{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Name_Expr,'tokens');% sig name
                           All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
                           All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
                           All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
                           All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
                           All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
                           All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                           u = u+1;
                       end
                end
           end
end
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Message %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Message_ID_Expr = 'BO_[\s]([0-9.]{1,})'
% Message_ID_Reg = regexp(Node_Sig_Reg{1,1},Message_ID_Expr,'tokens');  
% Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
% Message_Name_Reg = regexp(Node_Sig_Reg{1,1},Message_Name_Expr,'tokens'); 
% Message_Node_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
% Message_Node_Reg = regexp(Node_Sig_Reg{1,1},Message_Node_Expr,'tokens'); 
% 
% 
% 
% % %%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Signal %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% % Sig_Name = 
% Sig_Name_Expr = '[\s]{1,}SG_[\s]([a-z_A-Z_0-9_]){1,}';
% Sig_Name_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Name_Expr,'tokens'); 
% % Sig_Start_Value=  
% %  SG_ Eng_ActiveBrakLampCmd : 8|2@1+ (1,0) [0|3] "" Vector__XXX
% Sig_Start_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s]([-0-9.]){1,}[|]';
% Sig_Start_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Start_Expr,'tokens');
%��ȡ�����źŵĿ�ʼλ�� 

% Sig_Length_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|]([-0-9.]){1,}';
% Sig_Factor_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(]([-0-9.]){1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Bais_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,]([-0-9.]){1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Min_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[]([-0-9.]){1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Max_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|]([-0-9.]){1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% % Sig_Length_Value =
% Sig_Length_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|]([-0-9.]){1,}';
% Sig_Length_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Length_Expr,'tokens');
% % Sig_Factor_Value = 
% Sig_Factor_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(]([-0-9.]){1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Factor_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Factor_Expr,'tokens');
% % Sig_Bais_Value =
% Sig_Bais_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,]([-0-9.]){1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Bais_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Bais_Expr,'tokens');
% % Sig_Min_Value = 
% Sig_Min_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[]([-0-9.]){1,}[\|][-0-9.]{1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Min_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Min_Expr,'tokens');
% % Sig_Max_Value =
% Sig_Max_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|]([-0-9.]){1,}[\]][\s]["]["][\s][a-z_A-Z_0-9_]{1,}';
% Sig_Max_Reg = regexp(Node_Sig_Reg{1,1}{1,2},Sig_Max_Expr,'tokens');

% 
%% �����ݴ���excel��

try
    Excel = actxGetRunningServerr('Excel.application');
catch
    Excel = actxserver('Excel.application');
end
Excel.visible = 1;
Workbook = Excel.Workbooks.Add;
Sheet1 = Workbook.Sheet.Item(1);
Sheet2 = Workbook.Sheet.Item(2);
Sheet3 = Workbook.Sheet.Item(3);
Sheet1.Activate;
Sheet2.Activate;
Sheet3.Activate;
Sheet1.Name = '����';
Sheet2.Name = '�����б�';
Sheet3.Name = '���Խ��';
% ExcelActiveWorkbooks = get(Excel,'ActiveWorkBook');
% eSheet = get(ExcelActiveWorkbooks,'Sheets'); %��ǰ��sheet
% % eSheet.Add
% Sheet1 = Item(eSheet,1);
% Sheet2 = Item(eSheet,2);
% % eSheet.Add
% Sheet3 = Item(eSheet,3);
% eSheet.Item(1).Name = '����';
% eSheet.Item(2).Name = '�����б�';
% eSheet.Item(3).Name = '���Խ��';

%% ������ʾ
c_r=255;
c_g=255;
c_b=255;
Sheet1.Range('A1:F42').Interior.Color=hex2dec([dec2hex(c_r,2) dec2hex(c_g,2) dec2hex(c_b,2)])
Sheet1.Range('A1:F42').Font.Name= 'Times New Roman'; 
Sheet1.Range('A1:F42').Font.Size= 12; %��С
Sheet1.Range('A4:A4').value = '��Ŀ��';
Sheet1.Range('A4:A5').HorizontalAlignment = -4131; %�����
Sheet1.Range('A4:A5').Font.Size= 12; %��С
Sheet1.Range('A4:A5').Font.Name = '���Ŀ���';
Sheet1.Range('A4:A5').Font.Bold = 1; % �Ӵ�
Prj_Name = PrjName
% Prj_Name= 'LE';% ͨ��UI��������
Sheet1.Range('B4:B4').value = Prj_Name; 
Sheet1.Range('B4:B5').HorizontalAlignment = -4131; %�����
Sheet1.Range('B4:B5').Font.Size= 12; %��С
Sheet1.Range('B4:B5').Font.Name = '���Ŀ���';
Sheet1.Range('B4:B5').Font.Bold = 1; % �Ӵ�

Sheet1.Range('A5:A5').value = '���ţ�';


Depart_Name= '��ؼ����о���';% ͨ��UI��������
Sheet1.Range('B5:B5').value = Depart_Name;
Sheet1.Range('B5:B5').HorizontalAlignment = -4131; %�����
Car_Type = CarType
% Car_Type = 'LE8T���綯��ʽ�ػ���';
Type=strcat(Car_Type,'��������������');
Sheet1.Range('A8:A8').value = Type;
Sheet1.Range('A8:F8').MergeCells = 1;
Sheet1.Range('A8:F8').Font.Size= 24; %��С
Sheet1.Range('A8:F8').Font.Name = '���Ŀ���';
Sheet1.Range('A8:F8').Font.Bold = 1; % �Ӵ�
Sheet1.Range('A8:F8').ColumnWidth= 25; %
Sheet1.Range('A8:F8').HorizontalAlignment = -4108; %����

Sheet1.Range('A9:F9').MergeCells = 1;
Sheet1.Range('A9:F9').value = 'ͨ��Э���DBC���Լ�¼';
Sheet1.Range('A9:F9').Font.Size= 24; %��С
Sheet1.Range('A9:F9').Font.Name= '���Ŀ���';
Sheet1.Range('A9:F9').Font.Bold = 1; % �Ӵ�
Sheet1.Range('A9:F9').ColumnWidth= 25; %����
Sheet1.Range('A9:F9').HorizontalAlignment = -4108; %����


Sheet1.Range('B29:B29').value = '��    ����';
Sheet1.Range('B30:B30').value = 'У    �ԣ�';
Sheet1.Range('B31:B31').value = '��    �ˣ�';
Sheet1.Range('B32:B32').value = '��    ׼��';
Sheet1.Range('B29:E32').Font.Name= '���Ŀ���';
Sheet1.Range('B29:E32').Font.Bold = 1; % �Ӵ�
Sheet1.Range('B29:E32').Font.Size= 12; %��С
Sheet1.Range('B29:E32').HorizontalAlignment = -4131; %�����

edit = EditPerson
% edit='����';% ͨ��UI��������
Sheet1.Range('C29:C29').value = edit;
edit2=' ';% ͨ��UI��������
Sheet1.Range('C30:C30').value = edit2;
edit3=' ';% ͨ��UI��������
Sheet1.Range('C31:C31').value = edit3;
edit4=' ';% ͨ��UI��������
Sheet1.Range('C32:C32').value = edit4;

Sheet1.Range('D29:D29').value = 'ʱ    �䣺';
Sheet1.Range('D30:D30').value = 'ʱ    �䣺';
Sheet1.Range('D31:D31').value = 'ʱ    �䣺';
Sheet1.Range('D32:D32').value = 'ʱ    �䣺';


time = Time  
% time = '2020.9.25';
Sheet1.Range('E29:E29').value = time;
Sheet1.Range('E30:E30').value = time;
Sheet1.Range('E31:E31').value = time;
Sheet1.Range('E32:E32').value = time;
%% �����б���ʾ
Sheet2.Range('A1').value = '���';
Sheet2.Range('A1').ColumnWidth= 20; %
Sheet2.Range('B1').value = '��������';
Sheet2.Range('B1').ColumnWidth= 20; %
Sheet2.Range('C1').value = '�ڵ�����';
Sheet2.Range('C1').ColumnWidth= 20; %
Sheet2.Range('D1').value = '��������';
Sheet2.Range('D1').ColumnWidth= 25; %
Sheet2.Range('E1').value = '��������';
Sheet2.Range('E1').ColumnWidth= 45; %
Sheet2.Range('F1').value = '��ע�������ͼ����������Ҫ˵�����';
Sheet2.Range('F1').ColumnWidth= 55; %
Sheet2.Range('A1:F10').Border.LineStyle = 1;
Sheet2.Range('A1:F1').Font.Size=  11;
Sheet2.Range('A1:F1').Font.Bold=  1;
% Sheet2.Range('A1:F1').Font.Size=  20;
c_r=184;
c_g=204;
c_b=228;
Sheet2.Range('A1:F1').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])
c_r=255;
c_g=255;
c_b=255;
Sheet2.Range('A2:F10').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])
Sheet2.Range('A1:F10').RowHeight= 20; %
%% ���Խ����ʾ
%ͨ��Э�顢DBC�ļ��������
text = 'ͨѶЭ�顢DBC�ļ��������'
Sheet3.Range('B2').value = text;
Sheet3.Range('B2:P2').MergeCells = 1;
Sheet3.Range('B2:P2').HorizontalAlignment = -4108; %����
Sheet3.Range('B2:P2').Font.Size= 16; %��С
Sheet3.Range('B2:P2').Font.Bold= 1; %��С
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B2:P2').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])


Sheet3.Range('B3').value =  'ͨѶЭ��汾��';
version = 'KW90T'
Sheet3.Range('C3').value = version;
Sheet3.Range('C3:P3').MergeCells = 1;
Sheet3.Range('C3:P3').HorizontalAlignment = -4108; %����


Sheet3.Range('B4').value = 'DBC�汾��';
Sheet3.Range('C4').value = 'V2.0';
Sheet3.Range('C4:P4').MergeCells = 1;
Sheet3.Range('C4:P4').HorizontalAlignment = -4108; %����

test_person = 'XXX'
Sheet3.Range('B5').value = '������Ա';
Sheet3.Range('C5').value = test_person;
Sheet3.Range('C5:P5').MergeCells = 1;
Sheet3.Range('C5:P5').HorizontalAlignment = -4108; %����

test_time = 'XXX-XX-XX'
Sheet3.Range('B6').value =  '����ʱ��';
Sheet3.Range('C5').value = test_time;
Sheet3.Range('C6:P6').MergeCells = 1;


test_obj = '����Ŀ��'
tt1 = '1 ��������ͨ��Э�����ݵ������޶����ԣ�'
tt2 = '2 ����ͨѶЭ���DBC�ļ��еĽڵ㡢�źš����ĵ����ơ�ID��ַ�����ڡ���ʼ���ȡ��źų��ȡ��źŷֱ��ʡ��ź�ƫ����һ���Ժ�׼ȷ�ԣ�'
tt3 = 'ͨѶЭ���Լ�����ͨѶЭ��Ƶ�DBC�ļ�'
Sheet3.Range('B7').value = test_obj;
Sheet3.Range('C7').value = tt1;
Sheet3.Range('C8').value = tt2;
Sheet3.Range('B7:B8').MergeCells = 1;

Sheet3.Range('C7:P7').MergeCells = 1;
Sheet3.Range('C7:P7').HorizontalAlignment = -4108; %����
Sheet3.Range('C8:P8').MergeCells = 1;
Sheet3.Range('C8:P8').HorizontalAlignment = -4108; %����
test_scope = '���Է�Χ'
Sheet3.Range('C9').value = tt3;
Sheet3.Range('B9').value = test_scope;
Sheet3.Range('C9:P9').MergeCells = 1;
Sheet3.Range('C9:P9').HorizontalAlignment = -4108; %����
Sheet3.Range('B2:P9').Border.LineStyle = 1;
%������

context = '��������'
t1='ͨѶЭ���Ƿ�������������VCU������нڵ����˽ṹ��'
t2 = 'ͨѶЭ���Ƿ���CAN�����ͨѶ���ʣ�'
t3 = 'DBC�ļ��Ƿ������ÿ�����ĵ�ͨѶ���ڣ�'
t4 = 'ͨѶЭ���DBC�Ľڵ������Ƿ���һ�������壿'
result = '���Խ��'
finish = '������'
Sheet3.Range('B12').value = finish;
Sheet3.Range('B12:B15').MergeCells = 1;

Sheet3.Range('B11').value = context;
Sheet3.Range('B11').value = context;

Sheet3.Range('B11:P11').HorizontalAlignment = -4108; %����
Sheet3.Range('B11:P11').Font.Size= 16; %��С
Sheet3.Range('B11:P11').Font.Bold= 1; %��С
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B11:P11').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])


Sheet3.Range('O11').value = result;
Sheet3.Range('B11:N11').MergeCells = 1;
Sheet3.Range('B11:N11').HorizontalAlignment = -4108; %����
Sheet3.Range('B11:N11').Font.Size= 16; %��С

Sheet3.Range('O11:P11').MergeCells = 1;
Sheet3.Range('O11:P11').HorizontalAlignment =  -4108; %����
Sheet3.Range('O12:P12').MergeCells = 1;
Sheet3.Range('O13:P13').MergeCells = 1;
Sheet3.Range('O14:P14').MergeCells = 1;
Sheet3.Range('O15:P15').MergeCells = 1;

Sheet3.Range('C12').value = t1;
Sheet3.Range('C13').value = t2;
Sheet3.Range('C14').value = t3;
Sheet3.Range('C15').value = t4;
Sheet3.Range('C12:N12').MergeCells = 1;
Sheet3.Range('C13:N13').MergeCells = 1;
Sheet3.Range('C14:N14').MergeCells = 1;
Sheet3.Range('C15:N15').MergeCells = 1;
Sheet3.Range('B11:P15').Border.LineStyle = 1;


%��ȷ��һ����
% ���ɱ�ͷ ��A3��ʼ

baise_x = 2;
baise_y = 17;
% Header =['�ڵ�����';'�ڵ������Ƿ�һ��';'��������';'���������Ƿ�һ��';'���������Ƿ�һ��';'�ź�����';'�ź�����';'�ź�����'];

temp_x = char('A'+baise_x-1)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�����';

temp_x = char('A'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ڵ�����';
% Sheet3.Range(temp).ColumnWidth= 20; %



temp_x = char('B'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ڵ������Ƿ�һ��';
% Sheet3.Range(temp).ColumnWidth= 20; %

temp_x = char('C'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ڵ������Ƿ�һ��';

temp_x = char('D'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '��������';

temp_x = char('E'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '���������Ƿ�һ��';

temp_x = char('F'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '����ID��ַ�Ƿ�һ��';

temp_x = char('G'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '���������Ƿ�һ�� ��λ��ms��';

temp_x = char('H'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ź�����';

temp_x = char('I'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ź������Ƿ�һ��';

temp_x = char('J'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ź���ʼλ���Ƿ�һ��';

temp_x = char('K'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�źų����Ƿ�һ��';

temp_x = char('L'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�źŷֱ����Ƿ�һ��';

temp_x = char('M'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ź���Сֵ�Ƿ�һ��';

temp_x = char('N'+baise_x)
temo_y = num2str(2+baise_y);
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '�ź����ֵ�Ƿ�һ��';

Sheet3.Range('B19:P19').HorizontalAlignment = -4108; %����
Sheet3.Range('B19:P19').Font.Size= 11; %��С
Sheet3.Range('B19:P19').Font.Bold= 1; %��С
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B19:P19').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])

%% �ź�д��
cur_sig_number = 0;
sig_range_start_x = 'H'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
% �ź�����д��
[~,msg_per_row] = size(All); %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}) %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end);
        for i=1:sig_row{sig_num}
            all_sig{i,1} = char(All{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig;
               Sheet3.Range(sig_range).Columnwidth = 30;%excel���
        end

    end
end
% �ź���ʼλ��д��
cur_sig_number = 0;
sig_range_start_x = 'I'+baise_x+1;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%                        All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%                        All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%                        All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%                        All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%                        All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Start_Bit) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
      [sig_per_col,~] = size(All_Sig_Start_Bit{1,msg_num}); %ĳһ�ڵ������е��ź�����
      for sig_num=1:sig_per_col
    % ��ǰ�ź�����

    %��ȡ�ź�1����
        [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Start_Bit{1,msg_num}{sig_num,2}{1,2}) ;
        sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
        sig_range_Mid = strcat(sig_range_start,sig_range_mid);
        cur_sig_number = cur_sig_number+sig_row{sig_num};
        sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
        sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    % sig_range_end = num2str(2+sig_row)
        sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
        sig_range=strcat(sig_range_Mid,sig_range_end)
            for i=1:sig_row{sig_num}
                all_sig_start{i,1} = char(All_Sig_Start_Bit{1,msg_num}{sig_num,2}{1,2}{1,i});
        %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
                 Sheet3.Range(sig_range).value = all_sig_start;
                   Sheet3.Range(sig_range).Columnwidth = 25;%excel���
            end

        end
end
%�ź������Ƿ�һ�� 

cur_sig_number = 0;
sig_range_start_x = 'J'+baise_x-1;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%  All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%  All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%  All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%  All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%  All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Length) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Length{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            %all_sig_start{i,1} = char(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = [];
               Sheet3.Range(sig_range).Columnwidth = 20;%excel���
        end

    end
end
%�źų���д��

cur_sig_number = 0;
sig_range_start_x = 'J'+baise_x+1;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%  All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%  All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%  All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%  All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%  All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Length) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Length{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
               Sheet3.Range(sig_range).Columnwidth = 20;%excel���
        end

    end
end

%�źŷֱ���д��
cur_sig_number = 0;
sig_range_start_x = 'K'+baise_x+1;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%  All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%  All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%  All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%  All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%  All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Factor) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Factor{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Factor{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Factor{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
               Sheet3.Range(sig_range).Columnwidth = 30;%excel���
        end

    end
end
%�ź���Сֵд��
cur_sig_number = 0;
sig_range_start_x = 'L'+baise_x+1;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%                        All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%                        All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%                        All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%                        All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%                        All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Bais) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Bais{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Bais{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Bais{1,msg_num}{sig_num,2}{1,2}{1,i})
            Sheet3.Range(sig_range).value = all_sig_start;
            Sheet3.Range(sig_range).Columnwidth = 30;%excel���
        end
    end
end

%�ź����ֵд��
cur_sig_number = 0;
sig_range_start_x = 'M'+baise_x+1;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
%  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
%                        All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
%                        All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
%                        All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
%                        All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
%                        All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
                       
[~,msg_per_row] = size(All_Sig_Max) %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Max{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Max{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
             all_sig_start{i,1} = char(All_Sig_Max{1,msg_num}{sig_num,2}{1,2}{1,i});
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
             Sheet3.Range(sig_range).Columnwidth = 30;%excel���
        end
    end
end
% %�ź���Сֵд��
% 
% cur_sig_number = 0;
% sig_range_start_x = 'N'+baise_x+1;
% sig_range_end_x =sig_range_start_x;
% sig_range_mid = ':';
% sig_range_start_y = 3+baise_y;
% %  All_Sig_Start_Bit{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Start_Expr,'tokens');% sig name
% %                        All_Sig_Length{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Length_Expr,'tokens');% sig name
% %                        All_Sig_Factor{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Factor_Expr,'tokens');% sig name
% %                        All_Sig_Bais{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Bais_Expr,'tokens');% sig name
% %                        All_Sig_Min{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Min_Expr,'tokens');% sig name
% %                        All_Sig_Max{1,i}{new,2}{u,2} = regexp(Node_Sig_Reg{1,j}{1,2},Sig_Max_Expr,'tokens');% sig name
%                        
% [msg_per_col,msg_per_row] = size(All_Sig_Min) %ĳһ�ڵ������е��ź�����
% 
% for msg_num=1:msg_per_row 
% [sig_per_col,sig_per_row] = size(All_Sig_Min{1,msg_num}); %ĳһ�ڵ������е��ź�����
% 
% for sig_num=1:sig_per_col
% % ��ǰ�ź�����
% 
% %��ȡ�ź�1����
%     [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Min{1,msg_num}{sig_num,2}{1,2}) 
%     sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
%     sig_range_Mid = strcat(sig_range_start,sig_range_mid);
%     cur_sig_number = cur_sig_number+sig_row{sig_num};
%     sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
%     sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% % sig_range_end = num2str(2+sig_row)
%     sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
%     sig_range=strcat(sig_range_Mid,sig_range_end);
%         for i=1:sig_row{sig_num}
%             all_sig_start{i,1} = char(All_Sig_Min{1,msg_num}{sig_num,2}{1,2}{1,i})
%     %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
%              Sheet3.Range(sig_range).value = all_sig_start;
%                Sheet3.Range(sig_range).Columnwidth = 10;%excel���
%         end
%     end
% end
%�źų�ʼֵд��

%% ��������д��

cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'D'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
    for i=1:sig_w_row{sig_num}
        all_sig{i,1} = char(All{1,msg_num}{sig_num,2}{1,1})
        s_num =    s_num +    1;    %��ǰ�����ź�����
        all_msg{s_num,1}{i,1} = all_sig{i,1};
        [cc ,~]= size(all_msg{s_num,1})%
          if(cc~=1)   
             all_msg{s_num,1}=[];
          end
    end
end
end

sig_start_x = sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
[sig_mid_number, ~ ]= size(all_msg)
 sig_end_y = 3+sig_mid_number-1+baise_y;   %�������λ��
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end)
 for i = 1:size(all_msg)
        sig{i,1}=char(all_msg{i,1})
     
 end
 Sheet3.Range(C_range).value = sig;%���ձ�������д��

 %�����ĺϲ�
 %�жϺϲ�����
 Merger_times = 0; %�ϲ�����
 [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1})
     [Mer_col_last,~]  = size(sig{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


  %��¼�ϲ���ʼλ��
Merger_times_start  = 0;
 [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1})
     [Mer_col_last,~]  = size(sig{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
          Merger_times_start = Merger_times_start +1 ;
        start_situt{Merger_times_start,1} = i ; 
     end
 end
 Merger_times_end = 0;
 %��¼�ϲ�����λ��
  [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1})
     [Mer_col_last,~]  = size(sig{i+1,1})
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times_end = Merger_times_end +1 ;
        end_situt{Merger_times_end,1} = i ; 
     end
 end
 %�����һ������ֻ��һ�����ģ�����1
 [allmsg_c,~] = size(all_msg)
 [amsg_c,~]=size(all_msg{allmsg_c,1})
 if(amsg_c == 0)
     Merger_times_end = Merger_times_end + 1; %�������Ŀո�
     [Mer_col,~] = size(all_msg) 
     end_situt{Merger_times_end,1} =  Mer_col; 
 end
 %�ϲ���Χ

 for i =1:Merger_times
    sig__y = 3+start_situt{i,1}-1+baise_y;
    sig__start = strcat(sig_start_x,num2str(sig__y));% H3
    sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %���½���λ��
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed);
 end
 %��ʼ�ϲ�
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
       Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel���
 end
 

%���� id д��
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'F'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) ;
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end);
    for i=1:sig_w_row{sig_num}
           all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1})
           s_num  = s_num + 1;    %��ǰ�����ź�����
           all_msg{s_num,1}{i,1} = all_sig{i,1};
           [cc ,~] = size(all_msg{s_num,1});
          if(cc~=1)   
             all_msg{s_num,1}=[];
          end
    end
end
end

sig_start_x = sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
[sig_mid_number, ~ ]= size(all_msg);
 sig_end_y = 3+sig_mid_number-1+baise_y;   %�������λ��
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end);
 for i = 1:size(all_msg)
        sig{i,1} = char(all_msg{i,1});
     
 end
 Sheet3.Range(C_range).value = sig;%���ձ�������д��

 %�����ĺϲ�
 %�жϺϲ�����
 Merger_times = 0; %�ϲ�����
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


%��¼�ϲ���ʼλ��
Merger_times  = 0;
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
          Merger_times = Merger_times +1 ;
        start_situt{Merger_times,1} = i ; 
     end
 end
 Merger_times = 0;
 %��¼�ϲ�����λ��
  [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times = Merger_times +1 ;
        end_situt{Merger_times,1} = i ; 
     end
 end
 %�����һ������ֻ��һ�����ģ�����1
 [allmsg_c,~] = size(all_msg);
 [amsg_c,~]=size(all_msg{allmsg_c,1});
 if(amsg_c == 0)
     Merger_times = Merger_times + 1; %�������Ŀո�
     [Mer_col,~] = size(all_msg) ;
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %�ϲ���Χ

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %���½���λ��
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed)
 end
 %��ʼ�ϲ�
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 25; %excel���
 end
%
 
 
% �ź�����1
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'E'+baise_x;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) ;
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end)
    for i=1:sig_w_row{sig_num}
        all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1});
        s_num =    s_num +    1;    %��ǰ�����ź�����
        all_msg{s_num,1}{i,1} = all_sig{i,1};
        [cc ,~]= size(all_msg{s_num,1});
          if(cc~=1)   
             all_msg{s_num,1}=[];
          end
    end
end
end

sig_start_x = sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
[sig_mid_number, ~ ]= size(all_msg);
 sig_end_y = 3+sig_mid_number-1+baise_y;   %�������λ��
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end)
 for i = 1:size(all_msg)
        sig{i,1}=char(all_msg{i,1});
     
 end
 Sheet3.Range(C_range).value = [];%���ձ�������д��

 %�����ĺϲ�
 %�жϺϲ�����
 Merger_times = 0; %�ϲ�����
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


  %��¼�ϲ���ʼλ��
Merger_times  = 0;
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
          Merger_times = Merger_times +1 ;
        start_situt{Merger_times,1} = i ; 
     end
 end
 Merger_times = 0;
 %��¼�ϲ�����λ��
  [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times = Merger_times +1 ;
        end_situt{Merger_times,1} = i ; 
     end
 end
 [allmsg_c,~] = size(all_msg)
 [amsg_c,~]=size(all_msg{allmsg_c,1})
 if(amsg_c == 0)
     Merger_times = Merger_times + 1; %�������Ŀո�
     [Mer_col,~] = size(all_msg) 
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %�ϲ���Χ

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %���½���λ��
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed)
 end
 %��ʼ�ϲ�
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel���
 end

% ��������
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'G'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %ĳһ�ڵ������е��ź�����

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %ĳһ�ڵ������е��ź�����

for sig_num=1:sig_per_col
% ��ǰ�ź�����

%��ȡ�ź�1����
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
    sig_range=strcat(sig_range_Mid,sig_range_end);
    for i=1:sig_w_row{sig_num}
           all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1})
           s_num  = s_num + 1;    %��ǰ�����ź�����
           all_msg{s_num,1}{i,1} = all_sig{i,1};
           [cc ,~] = size(all_msg{s_num,1});%
          if(cc~=1)   
             all_msg{s_num,1}=[];
          end
    end
end
end

sig_start_x = sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
[sig_mid_number, ~ ]= size(all_msg);
 sig_end_y = 3+sig_mid_number-1+baise_y;   %�������λ��
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end);
 for i = 1:size(all_msg)
        sig{i,1} = char(all_msg{i,1});
 end
 %������id�滻�ɱ�������
[~,cycle_r] = size(Node_msg_cycle_time);
for i = 1:size(all_msg)
    for j = 1:cycle_r
        if(strcmp(sig{i,1},Node_msg_cycle_time{1,j}{1,1}))
            sig{i,1} = Node_msg_cycle_time{1,j}{1,2};
            break;
%         else
%            % sig{i,1} = [];
%             break;
        end
    end
end
% �޳���idһ�µ�����
% �����
%
 Sheet3.Range(C_range).value = sig;%���ձ�������д��

 %�����ĺϲ�
 %�жϺϲ�����
 Merger_times = 0; %�ϲ�����
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


%��¼�ϲ���ʼλ��
Merger_times  = 0;
 [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
          Merger_times = Merger_times +1 ;
        start_situt{Merger_times,1} = i ; 
     end
 end
 Merger_times = 0;
 %��¼�ϲ�����λ��
  [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times = Merger_times +1 ;
        end_situt{Merger_times,1} = i ; 
     end
 end
 %�����һ������ֻ��һ�����ģ�����1
 [allmsg_c,~] = size(all_msg);
 [amsg_c,~]=size(all_msg{allmsg_c,1});
 if(amsg_c == 0)
     Merger_times = Merger_times + 1; %�������Ŀո�
     [Mer_col,~] = size(all_msg) ;
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %�ϲ���Χ

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %���½���λ��
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed);
 end
 %��ʼ�ϲ�
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel���
 end
%
 

%% �ڵ�д��
%  
% cur_sig_number_node= 0;
all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'A'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %ĳһ�ڵ����� msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}); %ĳһ�ڵ������еı������� sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%ĳһ���Ľڵ����źŵ�����sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %����ýڵ��µ���������
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %����sig����
    sig_range_end_y = 3+sig_number-1+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end);
    all_all{sig_range_start_y-baise_y-2,1} = char(All{1,node}{1,1});
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
end
[null_num_c,~ ]= size(all_all);
 for m=1:cur_sig_number_node{msg_per_row}-2
              all_all{ null_num_c+1+m,1} = [];
end
sig_start_x =sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
sig_end = strcat(sig_end_x,num2str(sig_end_y));
range = strcat(sig_Mid,sig_end)
Sheet3.Range(range).value = all_all;

% ���ڵ�ϲ�

  %�жϺϲ�����
 Merger_node_times = 0; %�ϲ�����
 [Merger_node_c ,~] = size(all_all);
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1});
     [Mer_col_last,~]  = size(all_all{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%��¼�ϲ���ʼλ��
Merger_node_times  = 0;
 [sig_node_c ,~] = size(all_all)
 for i=1:sig_node_c-1
     [Mer_node_col_first,~] = size(all_all{i,1});
     [Mer_node_col_last,~]  = size(all_all{i+1,1});
     if(Mer_node_col_first == 1  &&  Mer_node_col_last == 0)
            Merger_node_times = Merger_node_times +1 ;
            start_node_situt{Merger_node_times,1} = i ; 
     end
 end
 
Merger_node_times = 0;
 %��¼�ϲ�����λ��
  [sig_Node_c ,~] = size(all_all);
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1});
     [Mer_c_last,~]  = size(all_all{i+1,1});
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %�������Ŀո�
 [Mer_n_col,~] = size(all_all) ;
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %�ϲ���Χ

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %���½���λ��
    sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
    Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %��ʼ�ϲ�
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 30;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/ͨ�ø�ʽ';
 end

% �ڵ�����1 %
all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'B'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %ĳһ�ڵ����� msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}) %ĳһ�ڵ������еı������� sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%ĳһ���Ľڵ����źŵ�����sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %����ýڵ��µ���������
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %����sig����
    sig_range_end_y = 3+sig_number-1+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end)
    all_all{sig_range_start_y-2-baise_y,1} = char(All{1,node}{1,1})
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
end

sig_start_x =sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
sig_end_y = 3+sig_number-1+baise_y;   %�������λ��
sig_end = strcat(sig_end_x,num2str(sig_end_y));
range = strcat(sig_Mid,sig_end)
Sheet3.Range(range).value = [];

% ���ڵ�ϲ�

  %�жϺϲ�����
 Merger_node_times = 0; %�ϲ�����
 [Merger_node_c ,~] = size(all_all);
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1})
     [Mer_col_last,~]  = size(all_all{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%��¼�ϲ���ʼλ��
Merger_node_times  = 0;
 [sig_node_c ,~] = size(all_all);
 for i=1:sig_node_c-1
     [Mer_node_col_first,~] = size(all_all{i,1})
     [Mer_node_col_last,~]  = size(all_all{i+1,1})
     if(Mer_node_col_first == 1  &&  Mer_node_col_last == 0)
            Merger_node_times = Merger_node_times +1 ;
            start_node_situt{Merger_node_times,1} = i ; 
     end
 end
 
Merger_node_times = 0;
 %��¼�ϲ�����λ��
  [sig_Node_c ,~] = size(all_all)
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1})
     [Mer_c_last,~]  = size(all_all{i+1,1})
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %�������Ŀո�
 [Mer_n_col,~] = size(all_all) 
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %�ϲ���Χ

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %���½���λ��
    sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
    Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %��ʼ�ϲ�
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 20;
%     Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/ͨ�ø�ʽ';
 end
%  
% % �ڵ�����2 %

all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'C'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %ĳһ�ڵ����� msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}) %ĳһ�ڵ������еı������� sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%ĳһ���Ľڵ����źŵ�����sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %����ýڵ��µ���������
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %����sig����
    sig_range_end_y = 3+sig_number-1+baise_y;   %�������λ��
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end)
    all_all{sig_range_start_y-2-baise_y,1} = char(All{1,node}{1,1})
    sig_range_start_y = sig_range_end_y + 1;%���¿�ʼy��λ��
end


sig_start_x =sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
%  sig_number = sig_number+cur_sig_number_node{node} ; %����sig����
 sig_end_y = 3+sig_number-1+baise_y;   %�������λ��
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
range = strcat(sig_Mid,sig_end)
Sheet3.Range(range).value = [];

% ���ڵ�ϲ�

  %�жϺϲ�����
 Merger_node_times = 0; %�ϲ�����
 [Merger_node_c ,~] = size(all_all)
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1})
     [Mer_col_last,~]  = size(all_all{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%��¼�ϲ���ʼλ��
Merger_node_times  = 0;
 [sig_node_c ,~] = size(all_all)
 for i=1:sig_node_c-1
     [Mer_node_col_first,~] = size(all_all{i,1})
     [Mer_node_col_last,~]  = size(all_all{i+1,1})
     if(Mer_node_col_first == 1  &&  Mer_node_col_last == 0)
            Merger_node_times = Merger_node_times +1 ;
            start_node_situt{Merger_node_times,1} = i ; 
     end
 end
 
Merger_node_times = 0;
 %��¼�ϲ�����λ��
  [sig_Node_c ,~] = size(all_all)
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1})
     [Mer_c_last,~]  = size(all_all{i+1,1})
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %�������Ŀո�
 [Mer_n_col,~] = size(all_all) 
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %�ϲ���Χ

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    %  sig_number = sig_number+cur_sig_number_node{node} ; %����sig����
%     [sig_mid_number, uc ]= size(all_msg)
     sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %���½���λ��
     sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
     Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %��ʼ�ϲ�
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 20;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/ͨ�ø�ʽ';
 end
 
 %% ������
sig_range_start_x1 = 'A'+baise_x-1;
sig_range_end_x1 = char(sig_range_start_x1);
sig_range_mid = ':';
sig_range_start_y1 = num2str(3+baise_y);


sig_range_start_y2 = num2str(3+baise_y+sig_number-1);
Range = strcat(strcat(strcat(strcat(sig_range_end_x1,sig_range_start_y1),sig_range_mid),sig_range_end_x1),sig_range_start_y2)
Range1 = strcat(sig_range_end_x1,sig_range_start_y1)
Sheet3.Range(Range1).value = 'һ����';
Sheet3.Range(Range).MergeCells = 1;
Sheet3.Range(Range).Columnwidth =10;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
Sheet3.Range(Range).Numberformat = 'G/ͨ�ø�ʽ';

sig_range_start_y1 = num2str(3+baise_y-1);
sig_range_end_x2 = char('O'+baise_x-1);
Range = strcat(strcat(strcat(strcat(sig_range_end_x1,sig_range_start_y1),sig_range_mid),sig_range_end_x2),sig_range_start_y2)
Sheet3.Range(Range).Border.LineStyle = 1;
 
 
 
 
 
 
 

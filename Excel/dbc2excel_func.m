%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                 %           %     %        % % % %
                 %           %     %        %
                 %           % % % %        % % % %
                 %           %     %        %
                 % % %       %     %        %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 功能：完成DBC文件读取生成EXCEL
% 作者：李皓飞 
% 时间：2020.3.10
% 编译环境：Matlab2016b
% 版本：Version 1.1
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 输入DBC要求:DBC信号单位只能包含a-z_A-Z_0-9/°Ω℃%.*或者为无单位
function dbc2excel_func(Path_Name,PrjName,CarType,EditPerson,Time)
%clear all ;
% clc;
%% 读取DBC文件信息
%file = fopen('demo_my.dbc','r');
file = fopen(Path_Name,'r');
%% 读取DBC文件信息
%file = fopen('demo_my.dbc','r');
text = fscanf(file,'%c');
fclose(file);
Node_Expr = 'BU_:([\s][a-z_A-Z_0-9_]{1,}){1,}';                   %node normalizing
Node_Reg = regexp(text,Node_Expr,'tokens');               %match*.XLS file
NodeExp='([a-z_A-Z_0-9_]{1,})[\s]{0,}';
Node = regexp(Node_Reg{1},NodeExp,'tokens');   % every node 
[~,number_node]= size(Node{1}); %node number 

% statistic number of node,message and signal
Node_Sig_Expr = '(BO_[\s][-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,})([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%.*_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}){1,}';
% Node_Sig_Expr = '(BO_[\s][-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,}:[\s]{1,}[-0-9.]{1,}[\s]{1,}[a-z_A-Z_0-9_]{1,})([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][\W\w]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}){1,}';
Node_Sig_Reg = regexp(text,Node_Sig_Expr,'tokens');  
[~,number_message] = size(Node_Sig_Reg); %node number 
Sig_Expr = '([\s]{0,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
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
% 将全部节点放入all中
for i=1:number_node
    All{1,i}= Node{1,1}{1,i};
end

%分类报文信息
Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
Message_ID_Expr = 'BO_[\s]([0-9.]{1,})'
for j=1:number_message
    Message_Node{j} = regexp(Message_Array{j},Message_Node_Expr,'tokens') %报文上的节点
end
All_Message_ID=[];
    for i=1:number_node
        n=1;
        for j=1:number_message
             if  strcmp(char(Node{1,1}{1,i}),char(Message_Node{j}{1}))        %比较节点名称与报文上节点的名称是否相同 
%                   All{i}{n,2}= regexp(Message_Array{1,j},Message_Name_Expr,'tokens');%如果相同，将报文写入相对应的位置 即（报文上节点1的名称与节点1名称相同，则把报文名称归类到节点1下）            
                     All{1,i}{n,2}= regexp(Message_Array{1,j},Message_Name_Expr,'tokens');%如果相同，将报文写入相对应的位置 即（报文上节点1的名称与节点1名称相同，则把报文名称归类到节点1下）
                     All_Message_ID{1,i}{n,2}= regexp(Message_Array{1,j},Message_ID_Expr,'tokens');%如果相同，将报文写入相对应的位置 即（报文上节点1的名称与节点1名称相同，则把报文名称归类到节点1下）             
                     n=n+1;
            end
        end
    end
    
 %过虑没有报文的节点
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
 
 
  %过虑报文 
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
 [~ ,number_node] = size(All); %更新节点数量
 

 
    
%信号正则表达式

% Sig_Expr = '([\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}[\s]{0,}){1,}';
Message_Name_Expr = 'BO_[\s][0-9.]{1,}[\s]{1,}([a-z_A-Z_0-9_]{1,})'
Sig_Name_Expr = '[\s]{1,}SG_[\s]([a-z_A-Z_0-9_]){1,}';
Sig_Start_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s]([-0-9.]){1,}[|]';
Sig_Length_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|]([-0-9.]){1,}';
Sig_Factor_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(]([-0-9.]){1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Bais_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,]([-0-9.]){1,}[\)][\s][\[][-0-9.]{1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Min_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[]([-0-9.]){1,}[\|][-0-9.]{1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
Sig_Max_Expr = '[\s]{1,}SG_[\s][a-z_A-Z_0-9_]{1,}[\s]:[\s][-0-9.]{1,}[|][-0-9.]{1,}[@][-0-9.]{1,}[+][\s][\(][-0-9.]{1,}[\,][-0-9.]{1,}[\)][\s][\[][-0-9.]{1,}[\|]([-0-9.]){1,}[\]][\s]["][a-z_A-Z_0-9/°Ω℃%_]{0,}["][\s]{1,}[a-z_A-Z_0-9,_]{1,}';
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
%分类信号信息
for j=1:number_message
    mesg_name{j} = regexp(Node_Sig_Reg{1,j},Message_Name_Expr,'tokens'); % 找出所有报文
           for  i=1:number_node   
                 u =1;
                [d{i},r{i}] = size(All{1,i});  % 找到每个节点下报文的数量
                for new=1:d{i} %将新的报文名称与包含信号报文名称比较
                    if  strcmp(char(All{1,i}{new,2}{1}),char(mesg_name{j}{1}{1}))        %比较节点名称与报文上节点的名称是否相同 
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
%读取所有信号的开始位置 

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
%% 将数据存入excel中

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
Sheet1.Name = '封面';
Sheet2.Name = '问题列表';
Sheet3.Name = '测试结果';
% ExcelActiveWorkbooks = get(Excel,'ActiveWorkBook');
% eSheet = get(ExcelActiveWorkbooks,'Sheets'); %当前的sheet
% % eSheet.Add
% Sheet1 = Item(eSheet,1);
% Sheet2 = Item(eSheet,2);
% % eSheet.Add
% Sheet3 = Item(eSheet,3);
% eSheet.Item(1).Name = '封面';
% eSheet.Item(2).Name = '问题列表';
% eSheet.Item(3).Name = '测试结果';

%% 封面显示
c_r=255;
c_g=255;
c_b=255;
Sheet1.Range('A1:F42').Interior.Color=hex2dec([dec2hex(c_r,2) dec2hex(c_g,2) dec2hex(c_b,2)])
Sheet1.Range('A1:F42').Font.Name= 'Times New Roman'; 
Sheet1.Range('A1:F42').Font.Size= 12; %大小
Sheet1.Range('A4:A4').value = '项目：';
Sheet1.Range('A4:A5').HorizontalAlignment = -4131; %靠左边
Sheet1.Range('A4:A5').Font.Size= 12; %大小
Sheet1.Range('A4:A5').Font.Name = '华文楷体';
Sheet1.Range('A4:A5').Font.Bold = 1; % 加粗
Prj_Name = PrjName
% Prj_Name= 'LE';% 通过UI界面设置
Sheet1.Range('B4:B4').value = Prj_Name; 
Sheet1.Range('B4:B5').HorizontalAlignment = -4131; %靠左边
Sheet1.Range('B4:B5').Font.Size= 12; %大小
Sheet1.Range('B4:B5').Font.Name = '华文楷体';
Sheet1.Range('B4:B5').Font.Bold = 1; % 加粗

Sheet1.Range('A5:A5').value = '部门：';


Depart_Name= '电控技术研究所';% 通过UI界面设置
Sheet1.Range('B5:B5').value = Depart_Name;
Sheet1.Range('B5:B5').HorizontalAlignment = -4131; %靠左边
Car_Type = CarType
% Car_Type = 'LE8T纯电动厢式载货车';
Type=strcat(Car_Type,'车型整车控制器');
Sheet1.Range('A8:A8').value = Type;
Sheet1.Range('A8:F8').MergeCells = 1;
Sheet1.Range('A8:F8').Font.Size= 24; %大小
Sheet1.Range('A8:F8').Font.Name = '华文楷体';
Sheet1.Range('A8:F8').Font.Bold = 1; % 加粗
Sheet1.Range('A8:F8').ColumnWidth= 25; %
Sheet1.Range('A8:F8').HorizontalAlignment = -4108; %居中

Sheet1.Range('A9:F9').MergeCells = 1;
Sheet1.Range('A9:F9').value = '通信协议和DBC测试记录';
Sheet1.Range('A9:F9').Font.Size= 24; %大小
Sheet1.Range('A9:F9').Font.Name= '华文楷体';
Sheet1.Range('A9:F9').Font.Bold = 1; % 加粗
Sheet1.Range('A9:F9').ColumnWidth= 25; %居中
Sheet1.Range('A9:F9').HorizontalAlignment = -4108; %居中


Sheet1.Range('B29:B29').value = '编    辑：';
Sheet1.Range('B30:B30').value = '校    对：';
Sheet1.Range('B31:B31').value = '审    核：';
Sheet1.Range('B32:B32').value = '批    准：';
Sheet1.Range('B29:E32').Font.Name= '华文楷体';
Sheet1.Range('B29:E32').Font.Bold = 1; % 加粗
Sheet1.Range('B29:E32').Font.Size= 12; %大小
Sheet1.Range('B29:E32').HorizontalAlignment = -4131; %靠左边

edit = EditPerson
% edit='李皓飞';% 通过UI界面设置
Sheet1.Range('C29:C29').value = edit;
edit2=' ';% 通过UI界面设置
Sheet1.Range('C30:C30').value = edit2;
edit3=' ';% 通过UI界面设置
Sheet1.Range('C31:C31').value = edit3;
edit4=' ';% 通过UI界面设置
Sheet1.Range('C32:C32').value = edit4;

Sheet1.Range('D29:D29').value = '时    间：';
Sheet1.Range('D30:D30').value = '时    间：';
Sheet1.Range('D31:D31').value = '时    间：';
Sheet1.Range('D32:D32').value = '时    间：';


time = Time  
% time = '2020.9.25';
Sheet1.Range('E29:E29').value = time;
Sheet1.Range('E30:E30').value = time;
Sheet1.Range('E31:E31').value = time;
Sheet1.Range('E32:E32').value = time;
%% 问题列表显示
Sheet2.Range('A1').value = '序号';
Sheet2.Range('A1').ColumnWidth= 20; %
Sheet2.Range('B1').value = '网络名称';
Sheet2.Range('B1').ColumnWidth= 20; %
Sheet2.Range('C1').value = '节点名称';
Sheet2.Range('C1').ColumnWidth= 20; %
Sheet2.Range('D1').value = '报文名称';
Sheet2.Range('D1').ColumnWidth= 25; %
Sheet2.Range('E1').value = '问题描述';
Sheet2.Range('E1').ColumnWidth= 45; %
Sheet2.Range('F1').value = '备注（问题截图或者其他需要说明事项）';
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
%% 测试结果显示
%通信协议、DBC文件检测属性
text = '通讯协议、DBC文件检测属性'
Sheet3.Range('B2').value = text;
Sheet3.Range('B2:P2').MergeCells = 1;
Sheet3.Range('B2:P2').HorizontalAlignment = -4108; %居中
Sheet3.Range('B2:P2').Font.Size= 16; %大小
Sheet3.Range('B2:P2').Font.Bold= 1; %大小
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B2:P2').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])


Sheet3.Range('B3').value =  '通讯协议版本号';
version = 'KW90T'
Sheet3.Range('C3').value = version;
Sheet3.Range('C3:P3').MergeCells = 1;
Sheet3.Range('C3:P3').HorizontalAlignment = -4108; %居中


Sheet3.Range('B4').value = 'DBC版本号';
Sheet3.Range('C4').value = 'V2.0';
Sheet3.Range('C4:P4').MergeCells = 1;
Sheet3.Range('C4:P4').HorizontalAlignment = -4108; %居中

test_person = 'XXX'
Sheet3.Range('B5').value = '测试人员';
Sheet3.Range('C5').value = test_person;
Sheet3.Range('C5:P5').MergeCells = 1;
Sheet3.Range('C5:P5').HorizontalAlignment = -4108; %居中

test_time = 'XXX-XX-XX'
Sheet3.Range('B6').value =  '测试时间';
Sheet3.Range('C5').value = test_time;
Sheet3.Range('C6:P6').MergeCells = 1;


test_obj = '测试目标'
tt1 = '1 测试整车通信协议内容的完整无二义性：'
tt2 = '2 整车通讯协议和DBC文件中的节点、信号、报文的名称、ID地址、周期、起始长度、信号长度、信号分辨率、信号偏移量一致性和准确性；'
tt3 = '通讯协议以及根据通讯协设计的DBC文件'
Sheet3.Range('B7').value = test_obj;
Sheet3.Range('C7').value = tt1;
Sheet3.Range('C8').value = tt2;
Sheet3.Range('B7:B8').MergeCells = 1;

Sheet3.Range('C7:P7').MergeCells = 1;
Sheet3.Range('C7:P7').HorizontalAlignment = -4108; %居中
Sheet3.Range('C8:P8').MergeCells = 1;
Sheet3.Range('C8:P8').HorizontalAlignment = -4108; %居中
test_scope = '测试范围'
Sheet3.Range('C9').value = tt3;
Sheet3.Range('B9').value = test_scope;
Sheet3.Range('C9:P9').MergeCells = 1;
Sheet3.Range('C9:P9').HorizontalAlignment = -4108; %居中
Sheet3.Range('B2:P9').Border.LineStyle = 1;
%完整性

context = '测试内容'
t1='通讯协议是否完整定义了与VCU相关所有节点拓扑结构？'
t2 = '通讯协议是否定义CAN网络的通讯速率？'
t3 = 'DBC文件是否包含了每个报文的通讯周期？'
t4 = '通讯协议和DBC的节点名称是否定义一致无歧义？'
result = '测试结果'
finish = '完整性'
Sheet3.Range('B12').value = finish;
Sheet3.Range('B12:B15').MergeCells = 1;

Sheet3.Range('B11').value = context;
Sheet3.Range('B11').value = context;

Sheet3.Range('B11:P11').HorizontalAlignment = -4108; %居中
Sheet3.Range('B11:P11').Font.Size= 16; %大小
Sheet3.Range('B11:P11').Font.Bold= 1; %大小
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B11:P11').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])


Sheet3.Range('O11').value = result;
Sheet3.Range('B11:N11').MergeCells = 1;
Sheet3.Range('B11:N11').HorizontalAlignment = -4108; %居中
Sheet3.Range('B11:N11').Font.Size= 16; %大小

Sheet3.Range('O11:P11').MergeCells = 1;
Sheet3.Range('O11:P11').HorizontalAlignment =  -4108; %居中
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


%正确性一致性
% 生成表头 从A3开始

baise_x = 2;
baise_y = 17;
% Header =['节点名称';'节点名称是否一致';'报文名称';'报文名称是否一致';'报文周期是否一致';'信号名称';'信号名称';'信号名称'];

temp_x = char('A'+baise_x-1)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '检测项';

temp_x = char('A'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '节点名称';
% Sheet3.Range(temp).ColumnWidth= 20; %



temp_x = char('B'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '节点名称是否一致';
% Sheet3.Range(temp).ColumnWidth= 20; %

temp_x = char('C'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '节点拓扑是否一致';

temp_x = char('D'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '报文名称';

temp_x = char('E'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '报文名称是否一致';

temp_x = char('F'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '报文ID地址是否一致';

temp_x = char('G'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '报文周期是否一致 单位（ms）';

temp_x = char('H'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号名称';

temp_x = char('I'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号名称是否一致';

temp_x = char('J'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号起始位置是否一致';

temp_x = char('K'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号长度是否一致';

temp_x = char('L'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号分辨率是否一致';

temp_x = char('M'+baise_x)
temo_y = num2str(2+baise_y)
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号最小值是否一致';

temp_x = char('N'+baise_x)
temo_y = num2str(2+baise_y);
temp=strcat(temp_x,temo_y)
Sheet3.Range(temp).value = '信号最大值是否一致';

Sheet3.Range('B19:P19').HorizontalAlignment = -4108; %居中
Sheet3.Range('B19:P19').Font.Size= 11; %大小
Sheet3.Range('B19:P19').Font.Bold= 1; %大小
c_r=184;
c_g=204;
c_b=228;
Sheet3.Range('B19:P19').Interior.Color=hex2dec([dec2hex(c_b,2) dec2hex(c_g,2) dec2hex(c_r,2)])

%% 信号写入
cur_sig_number = 0;
sig_range_start_x = 'H'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
% 信号名称写入
[~,msg_per_row] = size(All); %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}) %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end);
        for i=1:sig_row{sig_num}
            all_sig{i,1} = char(All{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig;
               Sheet3.Range(sig_range).Columnwidth = 30;%excel宽度
        end

    end
end
% 信号起始位置写入
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
                       
[~,msg_per_row] = size(All_Sig_Start_Bit) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
      [sig_per_col,~] = size(All_Sig_Start_Bit{1,msg_num}); %某一节点下所有的信号数量
      for sig_num=1:sig_per_col
    % 当前信号数量

    %获取信号1数量
        [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Start_Bit{1,msg_num}{sig_num,2}{1,2}) ;
        sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
        sig_range_Mid = strcat(sig_range_start,sig_range_mid);
        cur_sig_number = cur_sig_number+sig_row{sig_num};
        sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
        sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    % sig_range_end = num2str(2+sig_row)
        sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
        sig_range=strcat(sig_range_Mid,sig_range_end)
            for i=1:sig_row{sig_num}
                all_sig_start{i,1} = char(All_Sig_Start_Bit{1,msg_num}{sig_num,2}{1,2}{1,i});
        %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
                 Sheet3.Range(sig_range).value = all_sig_start;
                   Sheet3.Range(sig_range).Columnwidth = 25;%excel宽度
            end

        end
end
%信号名称是否一致 

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
                       
[~,msg_per_row] = size(All_Sig_Length) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Length{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            %all_sig_start{i,1} = char(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = [];
               Sheet3.Range(sig_range).Columnwidth = 20;%excel宽度
        end

    end
end
%信号长度写入

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
                       
[~,msg_per_row] = size(All_Sig_Length) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Length{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Length{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
               Sheet3.Range(sig_range).Columnwidth = 20;%excel宽度
        end

    end
end

%信号分辨率写入
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
                       
[~,msg_per_row] = size(All_Sig_Factor) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Factor{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Factor{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Factor{1,msg_num}{sig_num,2}{1,2}{1,i})
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
               Sheet3.Range(sig_range).Columnwidth = 30;%excel宽度
        end

    end
end
%信号最小值写入
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
                       
[~,msg_per_row] = size(All_Sig_Bais) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Bais{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Bais{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
            all_sig_start{i,1} = char(All_Sig_Bais{1,msg_num}{sig_num,2}{1,2}{1,i})
            Sheet3.Range(sig_range).value = all_sig_start;
            Sheet3.Range(sig_range).Columnwidth = 30;%excel宽度
        end
    end
end

%信号最大值写入
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
                       
[~,msg_per_row] = size(All_Sig_Max) %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All_Sig_Max{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Max{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% sig_range_end = num2str(2+sig_row)
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
        for i=1:sig_row{sig_num}
             all_sig_start{i,1} = char(All_Sig_Max{1,msg_num}{sig_num,2}{1,2}{1,i});
    %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
             Sheet3.Range(sig_range).value = all_sig_start;
             Sheet3.Range(sig_range).Columnwidth = 30;%excel宽度
        end
    end
end
% %信号最小值写入
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
% [msg_per_col,msg_per_row] = size(All_Sig_Min) %某一节点下所有的信号数量
% 
% for msg_num=1:msg_per_row 
% [sig_per_col,sig_per_row] = size(All_Sig_Min{1,msg_num}); %某一节点下所有的信号数量
% 
% for sig_num=1:sig_per_col
% % 当前信号数量
% 
% %获取信号1数量
%     [sig_col{sig_num},sig_row{sig_num}] = size(All_Sig_Min{1,msg_num}{sig_num,2}{1,2}) 
%     sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
%     sig_range_Mid = strcat(sig_range_start,sig_range_mid);
%     cur_sig_number = cur_sig_number+sig_row{sig_num};
%     sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
%     sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
% % sig_range_end = num2str(2+sig_row)
%     sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
%     sig_range=strcat(sig_range_Mid,sig_range_end);
%         for i=1:sig_row{sig_num}
%             all_sig_start{i,1} = char(All_Sig_Min{1,msg_num}{sig_num,2}{1,2}{1,i})
%     %     Sheet3.Range('J3:J4').value = {all_sig{1,1};all_sig{2,1};};
%              Sheet3.Range(sig_range).value = all_sig_start;
%                Sheet3.Range(sig_range).Columnwidth = 10;%excel宽度
%         end
%     end
% end
%信号初始值写入

%% 报文名称写入

cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'D'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
    for i=1:sig_w_row{sig_num}
        all_sig{i,1} = char(All{1,msg_num}{sig_num,2}{1,1})
        s_num =    s_num +    1;    %当前计算信号数量
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
 sig_end_y = 3+sig_mid_number-1+baise_y;   %计算结束位置
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end)
 for i = 1:size(all_msg)
        sig{i,1}=char(all_msg{i,1})
     
 end
 Sheet3.Range(C_range).value = sig;%最终报文名称写入

 %将报文合并
 %判断合并次数
 Merger_times = 0; %合并次数
 [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1})
     [Mer_col_last,~]  = size(sig{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


  %记录合并开始位置
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
 %记录合并结束位置
  [sig_sig_c ,~] = size(sig)
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1})
     [Mer_col_last,~]  = size(sig{i+1,1})
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times_end = Merger_times_end +1 ;
        end_situt{Merger_times_end,1} = i ; 
     end
 end
 %若最后一个报文只有一个报文，不加1
 [allmsg_c,~] = size(all_msg)
 [amsg_c,~]=size(all_msg{allmsg_c,1})
 if(amsg_c == 0)
     Merger_times_end = Merger_times_end + 1; %加入最后的空格
     [Mer_col,~] = size(all_msg) 
     end_situt{Merger_times_end,1} =  Mer_col; 
 end
 %合并范围

 for i =1:Merger_times
    sig__y = 3+start_situt{i,1}-1+baise_y;
    sig__start = strcat(sig_start_x,num2str(sig__y));% H3
    sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %更新结束位置
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed);
 end
 %开始合并
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
       Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel宽度
 end
 

%报文 id 写入
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'F'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) ;
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end);
    for i=1:sig_w_row{sig_num}
           all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1})
           s_num  = s_num + 1;    %当前计算信号数量
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
 sig_end_y = 3+sig_mid_number-1+baise_y;   %计算结束位置
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end);
 for i = 1:size(all_msg)
        sig{i,1} = char(all_msg{i,1});
     
 end
 Sheet3.Range(C_range).value = sig;%最终报文名称写入

 %将报文合并
 %判断合并次数
 Merger_times = 0; %合并次数
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


%记录合并开始位置
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
 %记录合并结束位置
  [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times = Merger_times +1 ;
        end_situt{Merger_times,1} = i ; 
     end
 end
 %若最后一个报文只有一个报文，不加1
 [allmsg_c,~] = size(all_msg);
 [amsg_c,~]=size(all_msg{allmsg_c,1});
 if(amsg_c == 0)
     Merger_times = Merger_times + 1; %加入最后的空格
     [Mer_col,~] = size(all_msg) ;
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %合并范围

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %更新结束位置
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed)
 end
 %开始合并
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 25; %excel宽度
 end
%
 
 
% 信号属性1
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'E'+baise_x;
sig_range_end_x =sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) ;
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end)
    for i=1:sig_w_row{sig_num}
        all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1});
        s_num =    s_num +    1;    %当前计算信号数量
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
 sig_end_y = 3+sig_mid_number-1+baise_y;   %计算结束位置
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end)
 for i = 1:size(all_msg)
        sig{i,1}=char(all_msg{i,1});
     
 end
 Sheet3.Range(C_range).value = [];%最终报文名称写入

 %将报文合并
 %判断合并次数
 Merger_times = 0; %合并次数
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


  %记录合并开始位置
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
 %记录合并结束位置
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
     Merger_times = Merger_times + 1; %加入最后的空格
     [Mer_col,~] = size(all_msg) 
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %合并范围

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %更新结束位置
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed)
 end
 %开始合并
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel宽度
 end

% 报文周期
cur_sig_number = 0;
s_num =0 ;
sig_range_start_x = 'G'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All); %某一节点下所有的信号数量

for msg_num=1:msg_per_row 
[sig_per_col,~] = size(All{1,msg_num}); %某一节点下所有的信号数量

for sig_num=1:sig_per_col
% 当前信号数量

%获取信号1数量
    [sig_col{sig_num},sig_w_row{sig_num}] = size(All{1,msg_num}{sig_num,2}{1,2}) 
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);
    cur_sig_number = cur_sig_number+sig_w_row{sig_num};
    sig_range_end_y = 2+cur_sig_number+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
    sig_range=strcat(sig_range_Mid,sig_range_end);
    for i=1:sig_w_row{sig_num}
           all_sig{i,1} = char(All_Message_ID{1,msg_num}{sig_num,2}{1,1})
           s_num  = s_num + 1;    %当前计算信号数量
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
 sig_end_y = 3+sig_mid_number-1+baise_y;   %计算结束位置
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
 C_range = strcat(sig_Mid,sig_end);
 for i = 1:size(all_msg)
        sig{i,1} = char(all_msg{i,1});
 end
 %将报文id替换成报文周期
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
% 剔除与id一致的周期
% 待完成
%
 Sheet3.Range(C_range).value = sig;%最终报文周期写入

 %将报文合并
 %判断合并次数
 Merger_times = 0; %合并次数
 [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_times = Merger_times +1 ;
     end 
 end


%记录合并开始位置
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
 %记录合并结束位置
  [sig_sig_c ,~] = size(sig);
 for i=1:sig_sig_c-1
     [Mer_col_first,~] = size(sig{i,1});
     [Mer_col_last,~]  = size(sig{i+1,1});
     if(Mer_col_first == 0  &&  Mer_col_last == 1)
          Merger_times = Merger_times +1 ;
        end_situt{Merger_times,1} = i ; 
     end
 end
 %若最后一个报文只有一个报文，不加1
 [allmsg_c,~] = size(all_msg);
 [amsg_c,~]=size(all_msg{allmsg_c,1});
 if(amsg_c == 0)
     Merger_times = Merger_times + 1; %加入最后的空格
     [Mer_col,~] = size(all_msg) ;
     end_situt{Merger_times,1} =  Mer_col; 
 end
 %合并范围

 for i =1:Merger_times
     sig__y = 3+start_situt{i,1}-1+baise_y;
     sig__start = strcat(sig_start_x,num2str(sig__y));% H3
     sig_Md = strcat(sig__start,sig_mid);  % H3:
     sig_en_y = 3+end_situt{i,1} -1+baise_y;   %更新结束位置
     sig_ed = strcat(sig_end_x,num2str(sig_en_y));
     Merger_range{i,1} = strcat(sig_Md,sig_ed);
 end
 %开始合并
 for i= 1:Merger_times
      Sheet3.Range(Merger_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_range{i,1}).Columnwidth = 30; %excel宽度
 end
%
 

%% 节点写入
%  
% cur_sig_number_node= 0;
all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'A'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %某一节点数量 msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}); %某一节点下所有的报文数量 sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%某一报文节点下信号的数量sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %计算该节点下的数量个数
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %计算sig数量
    sig_range_end_y = 3+sig_number-1+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end);
    all_all{sig_range_start_y-baise_y-2,1} = char(All{1,node}{1,1});
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
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

% 将节点合并

  %判断合并次数
 Merger_node_times = 0; %合并次数
 [Merger_node_c ,~] = size(all_all);
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1});
     [Mer_col_last,~]  = size(all_all{i+1,1});
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%记录合并开始位置
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
 %记录合并结束位置
  [sig_Node_c ,~] = size(all_all);
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1});
     [Mer_c_last,~]  = size(all_all{i+1,1});
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %加入最后的空格
 [Mer_n_col,~] = size(all_all) ;
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %合并范围

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %更新结束位置
    sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
    Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %开始合并
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 30;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/通用格式';
 end

% 节点属性1 %
all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'B'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %某一节点数量 msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}) %某一节点下所有的报文数量 sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%某一报文节点下信号的数量sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %计算该节点下的数量个数
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %计算sig数量
    sig_range_end_y = 3+sig_number-1+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end)
    all_all{sig_range_start_y-2-baise_y,1} = char(All{1,node}{1,1})
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
end

sig_start_x =sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
sig_end_y = 3+sig_number-1+baise_y;   %计算结束位置
sig_end = strcat(sig_end_x,num2str(sig_end_y));
range = strcat(sig_Mid,sig_end)
Sheet3.Range(range).value = [];

% 将节点合并

  %判断合并次数
 Merger_node_times = 0; %合并次数
 [Merger_node_c ,~] = size(all_all);
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1})
     [Mer_col_last,~]  = size(all_all{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%记录合并开始位置
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
 %记录合并结束位置
  [sig_Node_c ,~] = size(all_all)
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1})
     [Mer_c_last,~]  = size(all_all{i+1,1})
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %加入最后的空格
 [Mer_n_col,~] = size(all_all) 
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %合并范围

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %更新结束位置
    sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
    Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %开始合并
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 20;
%     Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/通用格式';
 end
%  
% % 节点属性2 %

all_sig_number= 0;
sig_number = 0;
sig_range_start_x = 'C'+baise_x;
sig_range_end_x = sig_range_start_x;
sig_range_mid = ':';
sig_range_start_y = 3+baise_y;
[~,msg_per_row] = size(All) %某一节点数量 msg_per_row
for node=1:msg_per_row
cur_sig_number_node{node}=[0];
null_num{node} = [];
end
for node =1:msg_per_row

    [sig_per_col,~] = size(All{1,node}) %某一节点下所有的报文数量 sig_per_col
    sig_d_col{node}=sig_per_col;
    for i=1:sig_d_col{node}
        [sig_d_col{i},sig_s_row{i}] = size(All{1,node}{i,2}{1,2}) ;%某一报文节点下信号的数量sig_row
        cur_sig_number_node{node} = cur_sig_number_node{node}+sig_s_row{i}; %计算该节点下的数量个数
    end
    sig_range_start = strcat(sig_range_start_x,num2str(sig_range_start_y));% H3
    sig_range_Mid = strcat(sig_range_start,sig_range_mid);  % H3:
    sig_number = sig_number+cur_sig_number_node{node} ; %计算sig数量
    sig_range_end_y = 3+sig_number-1+baise_y;   %计算结束位置
    sig_range_end = strcat(sig_range_end_x,num2str(sig_range_end_y));
    sig_range = strcat(sig_range_Mid,sig_range_end)
    all_all{sig_range_start_y-2-baise_y,1} = char(All{1,node}{1,1})
    sig_range_start_y = sig_range_end_y + 1;%更新开始y的位置
end


sig_start_x =sig_range_start_x;
sig_end_x = sig_range_end_x;
sig_mid = ':';
sig_start_y = 3+baise_y;
sig_start = strcat(sig_start_x,num2str(sig_start_y));% H3
sig_Mid = strcat(sig_start,sig_mid);  % H3:
%  sig_number = sig_number+cur_sig_number_node{node} ; %计算sig数量
 sig_end_y = 3+sig_number-1+baise_y;   %计算结束位置
 sig_end = strcat(sig_end_x,num2str(sig_end_y));
range = strcat(sig_Mid,sig_end)
Sheet3.Range(range).value = [];

% 将节点合并

  %判断合并次数
 Merger_node_times = 0; %合并次数
 [Merger_node_c ,~] = size(all_all)
 for i=1:Merger_node_c-1
     [Mer_col_first,~] = size(all_all{i,1})
     [Mer_col_last,~]  = size(all_all{i+1,1})
     if(Mer_col_first == 1  &&  Mer_col_last == 0)
        Merger_node_times = Merger_node_times +1 ;
     end 
 end
%记录合并开始位置
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
 %记录合并结束位置
  [sig_Node_c ,~] = size(all_all)
 for i=1:sig_Node_c-1
     [Mer_c_first,~] = size(all_all{i,1})
     [Mer_c_last,~]  = size(all_all{i+1,1})
     if(Mer_c_first == 0  &&  Mer_c_last == 1)
          Merger_node_times = Merger_node_times +1 ;
        end_node_situt{Merger_node_times,1} = i ; 
     end
 end
 Merger_node_times = Merger_node_times + 1; %加入最后的空格
 [Mer_n_col,~] = size(all_all) 
 end_node_situt{Merger_node_times,1} =  Mer_n_col; 

 %合并范围

 for i =1:Merger_node_times
    sig_node_y = 3+start_node_situt{i,1}-1+baise_y;
    sig_node_start = strcat(sig_start_x,num2str(sig_node_y));% H3
    sig_Node_Md = strcat(sig_node_start,sig_mid);  % H3:
    %  sig_number = sig_number+cur_sig_number_node{node} ; %计算sig数量
%     [sig_mid_number, uc ]= size(all_msg)
     sig_node_end_y = 3+end_node_situt{i,1} -1+baise_y;   %更新结束位置
     sig_node_ed = strcat(sig_end_x,num2str(sig_node_end_y));
     Merger_node_range{i,1} = strcat(sig_Node_Md,sig_node_ed)
 end
 %开始合并
 for i= 1:Merger_node_times
      Sheet3.Range(Merger_node_range{i,1}).MergeCells = 1;
      Sheet3.Range(Merger_node_range{i,1}).Columnwidth = 20;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
      Sheet3.Range(Merger_node_range{i,1}).Numberformat = 'G/通用格式';
 end
 
 %% 最左列
sig_range_start_x1 = 'A'+baise_x-1;
sig_range_end_x1 = char(sig_range_start_x1);
sig_range_mid = ':';
sig_range_start_y1 = num2str(3+baise_y);


sig_range_start_y2 = num2str(3+baise_y+sig_number-1);
Range = strcat(strcat(strcat(strcat(sig_range_end_x1,sig_range_start_y1),sig_range_mid),sig_range_end_x1),sig_range_start_y2)
Range1 = strcat(sig_range_end_x1,sig_range_start_y1)
Sheet3.Range(Range1).value = '一致性';
Sheet3.Range(Range).MergeCells = 1;
Sheet3.Range(Range).Columnwidth =10;
%       Sheet3.Range(Merger_node_range{i,1}).HorizontalAlignment = 1;
Sheet3.Range(Range).Numberformat = 'G/通用格式';

sig_range_start_y1 = num2str(3+baise_y-1);
sig_range_end_x2 = char('O'+baise_x-1);
Range = strcat(strcat(strcat(strcat(sig_range_end_x1,sig_range_start_y1),sig_range_mid),sig_range_end_x2),sig_range_start_y2)
Sheet3.Range(Range).Border.LineStyle = 1;
 
 
 
 
 
 
 

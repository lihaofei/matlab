function varargout = context(varargin)
% CONTEXT MATLAB code for context.fig
%      CONTEXT, by itself, creates a new CONTEXT or raises the existing
%      singleton*.
%
%      H = CONTEXT returns the handle to a new CONTEXT or the handle to
%      the existing singleton*.
%
%      CONTEXT('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in CONTEXT.M with the given input arguments.
%
%      CONTEXT('Property','Value',...) creates a new CONTEXT or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before context_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to context_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help context

% Last Modified by GUIDE v2.5 25-Nov-2020 14:04:23

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @context_OpeningFcn, ...
                   'gui_OutputFcn',  @context_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
   
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before context is made visible.
function context_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to context (see VARARGIN)

% Choose default command line output for context
handles.output = hObject;
global id_num;
id_num = varargin{1}
id_num = id_num;
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes context wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = context_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



% --- Executes on button press in Btn_Ok.
function Btn_Ok_Callback(hObject, eventdata, handles)
% hObject    handle to Btn_Ok (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
            global id_num;
            global DeleteOrChangeFlag;
            db_name        = 'postgres';
            db_user_name   = 'postgres';
            db_user_pass   = 'postgre';
            db_server_ip   = 'localhost';
            db_server_port = '5432';
            url            = strcat('jdbc:postgresql://', db_server_ip, ':', db_server_port, '/');
            connection     = database(db_name, db_user_name, db_user_pass, 'org.postgresql.Driver', url);
            if (get(handles.cbox_Name,'Value'))
                option_name = 1;
            else
                option_name = 0;
            end
            if (get(handles.cbox_Age,'Value'))
                option_age = 1;
            else
                option_age = 0;
            end
            if (get(handles.cbox_Address,'Value'))
                option_address = 1;
            else
                option_address = 0;
            end
            if (get(handles.cbox_Salary,'Value'))
                option_salary = 1;
            else
                option_salary = 0;
            end
            if(DeleteOrChangeFlag==1)
                % 删除
                command11 = " DELETE FROM COMPANY WHERE ID = 0_XFXFX;";
                temp = id_num;
                temp = strrep(command11,'0_XFXFX',temp);
                command = temp;
                exec(connection,command);
            end
            
            if(DeleteOrChangeFlag == 2)
                % 更新名称
                 command12 =  "UPDATE COMPANY SET NAME = '0X_NAME' WHERE ID = 0X_ID;";
                 if(option_name==1)
                        name_temp =  get(handles.edit_Name,'String')
                        temp =id_num;
                        temp = strrep(command12,'0X_ID',temp);
                        temp = strrep(temp,'0X_NAME',name_temp);
                        command = temp;
                        exec(connection,command);
                 end
                % 更新年龄
                 command13 =  "UPDATE COMPANY SET AGE = 0X_AGE WHERE ID = 0X_ID;";
                 if(option_age==1)
                        age_temp =  get(handles.edit_Age,'String')
                        temp = id_num;
                        temp = strrep(command13,'0X_ID',temp);
                        temp = strrep(temp,'0X_AGE',age_temp);
                        command = temp;
                        exec(connection,command);
                 end
                % 更新地址
                  command14 =  "UPDATE COMPANY SET ADDRESS = '0X_ADDRESS' WHERE ID = 0X_ID;";
                 if(option_address==1)
                        address_temp =  get(handles.edit_Address,'String')
                        temp = id_num;
                        temp = strrep(command14,'0X_ID',temp);
                        temp = strrep(temp,'0X_ADDRESS',address_temp);
                        command = temp;
                        exec(connection,command);
                 end
                % 更新薪资
                  command15 =  "UPDATE COMPANY SET SALARY = 0X_SALARY WHERE ID = 0X_ID;";
                 if(option_salary==1)
                        salary_temp =  get(handles.edit_Salary,'String')
                        temp = id_num;
                        temp = strrep(command15,'0X_ID',temp);
                        temp = strrep(temp,'0X_SALARY',salary_temp);
                        command = temp;
                        exec(connection,command);
                 end
               
            end
            delete(handles.figure1)
            
%            delete(figure1);
%             close(context);
            % 更新
%             command12 =  "UPDATE COMPANY SET SALARY = 15000 WHERE ID = 3;";
%             % 修改
%             command13 =  "SELECT * FROM company;";
%             % 筛选出ID
%             command14 = "SELECT ID FROM company;";
% 
% msgbox('1')


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)



function edit_Name_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Name (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Name as text
%        str2double(get(hObject,'String')) returns contents of edit_Name as a double


% --- Executes during object creation, after setting all properties.
function edit_Name_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Name (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit3 as text
%        str2double(get(hObject,'String')) returns contents of edit3 as a double


% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit4 as text
%        str2double(get(hObject,'String')) returns contents of edit4 as a double


% --- Executes during object creation, after setting all properties.
function edit4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in rBtn_changed.
function rBtn_changed_Callback(hObject, eventdata, handles)
% hObject    handle to rBtn_changed (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global DeleteOrChangeFlag;
set(handles.rBtn_delete,'Value',0)
set(handles.rBtn_changed,'Value',1)
set(handles.edit_Name,'Visible','on')
set(handles.edit_Age,'Visible','on')
set(handles.edit_Address,'Visible','on')
set(handles.edit_Salary,'Visible','on')
set(handles.cbox_Name,'Visible','on')
set(handles.cbox_Age,'Visible','on')
set(handles.cbox_Address,'Visible','on')
set(handles.cbox_Salary,'Visible','on')
DeleteOrChangeFlag = 2;
% Hint: get(hObject,'Value') returns toggle state of rBtn_changed


% --- Executes on button press in rBtn_delete.
function rBtn_delete_Callback(hObject, eventdata, handles)
% hObject    handle to rBtn_delete (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global DeleteOrChangeFlag;
set(handles.rBtn_changed,'Value',0)
set(handles.edit_Name,'Visible','off')
set(handles.edit_Age,'Visible','off')
set(handles.edit_Address,'Visible','off')
set(handles.edit_Salary,'Visible','off')
set(handles.cbox_Name,'Visible','off')
set(handles.cbox_Age,'Visible','off')
set(handles.cbox_Address,'Visible','off')
set(handles.cbox_Salary,'Visible','off')
DeleteOrChangeFlag = 1;
% set(handles.rBtn_changed,'Value',1)
% Hint: get(hObject,'Value') returns toggle state of rBtn_delete


% --- Executes on button press in cbox_Name.
function cbox_Name_Callback(hObject, eventdata, handles)
% hObject    handle to cbox_Name (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of cbox_Name


% --- Executes on button press in cbox_Age.
function cbox_Age_Callback(hObject, eventdata, handles)
% hObject    handle to cbox_Age (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of cbox_Age


% --- Executes on button press in cbox_Address.
function cbox_Address_Callback(hObject, eventdata, handles)
% hObject    handle to cbox_Address (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of cbox_Address


% --- Executes on button press in cbox_Salary.
function cbox_Salary_Callback(hObject, eventdata, handles)
% hObject    handle to cbox_Salary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of cbox_Salary



function edit_Age_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Age (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Age as text
%        str2double(get(hObject,'String')) returns contents of edit_Age as a double


% --- Executes during object creation, after setting all properties.
function edit_Age_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Age (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_Address_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Address (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Address as text
%        str2double(get(hObject,'String')) returns contents of edit_Address as a double


% --- Executes during object creation, after setting all properties.
function edit_Address_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Address (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_Salary_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Salary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Salary as text
%        str2double(get(hObject,'String')) returns contents of edit_Salary as a double


% --- Executes during object creation, after setting all properties.
function edit_Salary_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Salary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function figure1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
% g = get(handles.rBtn_delete,'value')
% msgbox('2')
global DeleteOrChangeFlag;
DeleteOrChangeFlag= 1;

% handles.in1= varargin{1};
% guidata(hObject,handles);


% --- Executes when user attempts to close figure1.
function figure1_CloseRequestFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: delete(hObject) closes the figure
delete(hObject);

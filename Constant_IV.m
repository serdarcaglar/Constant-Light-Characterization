function varargout = Constant_IV(varargin)
% CONSTANT_IV MATLAB code for Constant_IV.fig
%      CONSTANT_IV, by itself, creates a new CONSTANT_IV or raises the existing
%      singleton*.
%
%      H = CONSTANT_IV returns the handle to a new CONSTANT_IV or the handle to
%      the existing singleton*.
%
%      CONSTANT_IV('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in CONSTANT_IV.M with the given input arguments.
%
%      CONSTANT_IV('Property','Value',...) creates a new CONSTANT_IV or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Constant_IV_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Constant_IV_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Constant_IV

% Last Modified by GUIDE v2.5 18-Sep-2020 16:38:18

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
    'gui_Singleton',  gui_Singleton, ...
    'gui_OpeningFcn', @Constant_IV_OpeningFcn, ...
    'gui_OutputFcn',  @Constant_IV_OutputFcn, ...
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


% --- Executes just before Constant_IV is made visible.
function Constant_IV_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Constant_IV (see VARARGIN)

% Choose default command line output for Constant_IV
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

clear all
global obj
% handles.obj = ardiuno();






% --- Outputs from this function are returned to the command line.
function varargout = Constant_IV_OutputFcn(hObject, eventdata, handles)
varargout{1} = handles.output;


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

contents=cellstr(get(hObject,'string'));
pop_choice=contents{get(hObject,'Value')};
global ser
handles.ser=serialport(pop_choice,1200);
ser=handles.ser;
guidata(hObject, handles);
try
    fopen(ser);
catch e
    msgbox('CommunicationERROR!!!')
    newobjs = instrfind;
    fclose(newobjs);
end

% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
freeports = serialportlist("available");
freeports=['---' freeports];
set(hObject,'String',freeports);


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
desktop=winqueryreg('HKEY_CURRENT_USER', 'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 'Desktop');
[file,path,~]=uiputfile([desktop,'\','constantlightexp.xlsx']);
A={'Date and Time','Voc','Isc'};
file_path = [path,file];
xlswrite(file_path,A)
devam =true;
sayac = 1;
while devam
    A(sayac,1) = datetime;
    A(sayac,2) = read_voltage;
    A(sayac,3) = read_current;
    if ~devam
        break;
    end
    sayac = sayac+1;
    xlswrite(file_path,A)
    puase(60)
end


function Voc = read_voltage()
 ser = handles.ser;
 obj = handles.obj;
 writeDigitalPin(obj, 'D8', 1);
 pause(0.5)
 writeline(ser,'VDC')
 pause(0.5)
 Voc = writeread('READ?');

function Isc= read_current()
 ser = handles.ser;
 obj = handles.obj;
 writeDigitalPin(obj, 'D8', 0);
 pause(0.5)
 writeline(ser,'ADC')
 pause(0.5)
 Isc = writeread('READ?');

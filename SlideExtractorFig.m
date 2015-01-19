function varargout = SlideExtractorFig(varargin)
% SLIDEEXTRACTORFIG MATLAB code for SlideExtractorFig.fig
%      SLIDEEXTRACTORFIG, by itself, creates a new SLIDEEXTRACTORFIG or raises the existing
%      singleton*.
%
%      H = SLIDEEXTRACTORFIG returns the handle to a new SLIDEEXTRACTORFIG or the handle to
%      the existing singleton*.
%
%      SLIDEEXTRACTORFIG('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in SLIDEEXTRACTORFIG.M with the given input arguments.
%
%      SLIDEEXTRACTORFIG('Property','Value',...) creates a new SLIDEEXTRACTORFIG or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before SlideExtractorFig_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to SlideExtractorFig_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help SlideExtractorFig

% Last Modified by GUIDE v2.5 18-Jan-2015 18:26:36

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @SlideExtractorFig_OpeningFcn, ...
                   'gui_OutputFcn',  @SlideExtractorFig_OutputFcn, ...
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


% --- Executes just before SlideExtractorFig is made visible.
function SlideExtractorFig_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to SlideExtractorFig (see VARARGIN)

% Choose default command line output for SlideExtractorFig
handles.output = hObject;

% Clear Edit Boxes
set(handles.e_ppt_folder_folder, 'String',' ');
set(handles.e_settings_slides_to_extract, 'String','[ 1 ]');
set(handles.e_settings_ppt_prefix, 'String',' ');

% Initialize Button
set(handles.pb_extract, 'Enable','off');

% Initialize Global Variables
handles.pptFolderPath = strcat(cd,'\');
handles.pptFolderName = '';
handles.pptFileNamePrefix = '';
handles.save_folder_name = strcat(cd,'\');
handles.pptFolderIsSelected = 0;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes SlideExtractorFig wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = SlideExtractorFig_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% redo but in pixel
% Set starting position in characters. Had problems with pixels
left_border = .8;
header = 2.0;
set(0,'Units','character') ;
screen_size = get(0,'ScreenSize');
set(handles.figure1,'Units','character');
dlg_size    = get(handles.figure1, 'Position');
pos1 = [ left_border , screen_size(4)-dlg_size(4)-1*header,...
    dlg_size(3) , dlg_size(4)];
set(handles.figure1,'Units','character');
set(handles.figure1,'Position',pos1);

% --- Executes on button press in pb_fig_folder.
function pb_fig_folder_Callback(hObject, eventdata, handles)
% hObject    handle to pb_fig_folder (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Open folder dialog box
save_folder_name = handles.save_folder_name;
dialog_title = 'PPT Save Directory';
save_folder_name = uigetdir(save_folder_name, dialog_title);

% Check return values
if isstr(save_folder_name)
   % user selected a folder
   handles.save_folder_name = strcat(save_folder_name,'\');
   
   % Update handles structure
    guidata(hObject, handles);
end


% --- Executes on button press in pb_fig_about.
function pb_fig_about_Callback(hObject, eventdata, handles)
% hObject    handle to pb_fig_about (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

SlideExtractorFigAbout;

% --- Executes on button press in pb_fig_ok.
function pb_fig_ok_Callback(hObject, eventdata, handles)
% hObject    handle to pb_fig_ok (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Close GUI
close SlideExtractorFig

% --- Executes on selection change in pm_settings_num_slides_ppt.
function pm_settings_num_slides_ppt_Callback(hObject, eventdata, handles)
% hObject    handle to pm_settings_num_slides_ppt (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns pm_settings_num_slides_ppt contents as cell array
%        contents{get(hObject,'Value')} returns selected item from pm_settings_num_slides_ppt


% --- Executes during object creation, after setting all properties.
function pm_settings_num_slides_ppt_CreateFcn(hObject, eventdata, handles)
% hObject    handle to pm_settings_num_slides_ppt (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function e_settings_slides_to_extract_Callback(hObject, eventdata, handles)
% hObject    handle to e_settings_slides_to_extract (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of e_settings_slides_to_extract as text
%        str2double(get(hObject,'String')) returns contents of e_settings_slides_to_extract as a double


% --- Executes during object creation, after setting all properties.
function e_settings_slides_to_extract_CreateFcn(hObject, eventdata, handles)
% hObject    handle to e_settings_slides_to_extract (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function e_settings_ppt_prefix_Callback(hObject, eventdata, handles)
% hObject    handle to e_settings_ppt_prefix (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of e_settings_ppt_prefix as text
%        str2double(get(hObject,'String')) returns contents of e_settings_ppt_prefix as a double


% --- Executes during object creation, after setting all properties.
function e_settings_ppt_prefix_CreateFcn(hObject, eventdata, handles)
% hObject    handle to e_settings_ppt_prefix (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function e_ppt_folder_folder_Callback(hObject, eventdata, handles)
% hObject    handle to e_ppt_folder_folder (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of e_ppt_folder_folder as text
%        str2double(get(hObject,'String')) returns contents of e_ppt_folder_folder as a double



% --- Executes during object creation, after setting all properties.
function e_ppt_folder_folder_CreateFcn(hObject, eventdata, handles)
% hObject    handle to e_ppt_folder_folder (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pb_ppt_folder_load.
function pb_ppt_folder_load_Callback(hObject, eventdata, handles)
% hObject    handle to pb_ppt_folder_load (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pb_ppt_folder_select_folder.
function pb_ppt_folder_select_folder_Callback(hObject, eventdata, handles)
% hObject    handle to pb_ppt_folder_select_folder (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% Open folder dialog box
pptFolderPath = handles.pptFolderPath;
dialog_title = 'PPT Folder Path';
user_selected_folder_path = uigetdir(pptFolderPath, dialog_title);

% Check return values
if isstr(user_selected_folder_path)
    % Get Folder Name
    separatorIndex = findstr(user_selected_folder_path, '\');
    pptFolderName = user_selected_folder_path(separatorIndex(end)+1:end);
    pptFileNamePrefix = strcat(pptFolderName, '_');
    
    % user selected a folder
    handles.pptFolderName = pptFolderName;
    handles.pptFileNamePrefix = pptFileNamePrefix;
    handles.pptFolderPath = strcat(user_selected_folder_path,'\');
   
    % Set Folder Name and PPT prefix name
    set(handles.e_ppt_folder_folder, 'String', pptFolderName);
    set(handles.e_settings_ppt_prefix, 'String', pptFileNamePrefix);
    
    % Allow Extract to begin
    handles.pptFolderIsSelected = 1;
    set(handles.pb_extract, 'Enable','on');
    
    % Update handles structure
    guidata(hObject, handles);
end


% --- Executes on button press in pb_extract.
function pb_extract_Callback(hObject, eventdata, handles)
% hObject    handle to pb_extract (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Check that folder has been selected
pptFolderIsSelected = handles.pptFolderIsSelected;

if pptFolderIsSelected == 1
    % Get SlideExtractor Parameters
    pptFolderPath = handles.pptFolderPath;
    pptFileNamePrefix = handles.pptFileNamePrefix;
    save_folder_name = handles.save_folder_name;
    
    % Create Extractor Class
    secObj = PptSlideExtractorClass...
        (pptFolderPath, pptFileNamePrefix, save_folder_name);
    
    % Get Extraction Properties
    slidesToExtract = ...
        eval(get(handles.e_settings_slides_to_extract, 'String'));
    numSlidesPerPPT = get(handles.pm_settings_num_slides_ppt, 'Value');
    
    % Begin slide extraction
    secObj = secObj.extractSlides (slidesToExtract, numSlidesPerPPT);
end

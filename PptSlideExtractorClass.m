classdef PptSlideExtractorClass
    %PptSlideExtractorClass Remove selected slides from each file in a PPT list
    %   
    % Input
    %
    %        pptFolderPath : Folder path that contains PPT/PPTX files
    %    pptFileNamePrefix : Destination PPT file name prefix
    %     save_folder_name : Destination Path to save destination PPT
    %      slidesToExtract : Array of slides to remove from each PPT
    %      numSlidesPerPPT : Number of slides to save in each destintation 
    %                        PPT. Each created PPT is numbered sequentially 
    %                        starting with 1.
    %
    % Function Prototypes
    %
    %    obj = PptSlideExtractorClass(pptFolderPath, pptFileNamePrefix, save_folder_name)
    %    obj = obj.extractSlides
    %    obj = obj.extractSlides(slidesToExtract, numSlidesPerPPT)
    %
    % Requirements
    %
    %    Tested with Microsoft PowerPoint 2010
    %
    % Acknowledgements
    %
    %    The following open source utilities are called by
    %    SlideExtractorClass
    %   
    %    saveppt2
    %    http://www.mathworks.com/matlabcentral/fileexchange/19322-saveppt2
    %
    %    dirr
    %    http://www.mathworks.com/matlabcentral/fileexchange/8682-dirr--find-files-recursively-filtering-name--date-or-bytes-
    %
    %
    %  Version: 1.0.1
    %
    % ---------------------------------------------
    % Dennis A. Dean, II, Ph.D
    %
    % Program for Sleep and Cardiovascular Medicine
    % Brigam and Women's Hospital
    % Harvard Medical School
    % 221 Longwood Ave
    % Boston, MA  02115
    %
    % File created: January 18, 2015
    % Last updated: January 19, 2015 
    %    
    % Copyright © [2014] The Brigham and Women's Hospital, Inc. THE BRIGHAM AND 
    % WOMEN'S HOSPITAL, INC. AND ITS AGENTS RETAIN ALL RIGHTS TO THIS SOFTWARE 
    % AND ARE MAKING THE SOFTWARE AVAILABLE ONLY FOR SCIENTIFIC RESEARCH 
    % PURPOSES. THE SOFTWARE SHALL NOT BE USED FOR ANY OTHER PURPOSES, AND IS
    % BEING MADE AVAILABLE WITHOUT WARRANTY OF ANY KIND, EXPRESSED OR IMPLIED, 
    % INCLUDING BUT NOT LIMITED TO IMPLIED WARRANTIES OF MERCHANTABILITY AND 
    % FITNESS FOR A PARTICULAR PURPOSE. THE BRIGHAM AND WOMEN'S HOSPITAL, INC. 
    % AND ITS AGENTS SHALL NOT BE LIABLE FOR ANY CLAIMS, LIABILITIES, OR LOSSES 
    % RELATING TO OR ARISING FROM ANY USE OF THIS SOFTWARE.
    %
    
    %---------------------------------------------------- Public Properties
    properties (Access = public)    
        % File/Path Parameters
        pptFolderPath
        pptFileNamePrefix
        save_folder_name
        
        % Extraction Parameters
        slidesToExtract = [1];
        numSlidesPerPPT = 100;
        
        % PPT Search String
        pptSearchStr = '\.ppt';
    end
    %------------------------------------------------- Dependent Properties
    properties (Dependent = true)
        
    end
    %--------------------------------------------------- Private Properties
    properties (Access = protected)
        % Operation parameters
        PARAMETERS_ARE_SET = 0;
        
        % PPT File Name Information
        filePathP
        fileNameP
        numPptFilesP
        pptFnWithPathP     
    end
    %------------------------------------------------------- Public Methods
    methods
        %------------------------------------------------------ Constructor
         function obj = PptSlideExtractorClass(varargin)
        
            if nargin == 3
                % Set input values
                obj.pptFolderPath = varargin{1};
                obj.pptFileNamePrefix = varargin{2};
                obj.save_folder_name = varargin{3};
                
                % Set Flag
                PARAMETERS_ARE_SET = 1;
            else
                % Echo function prototype to console
                fprintf('obj = SlideExtractorClass(pptFolderPath, pptFileNamePrefix, save_folder_name)');
            end
        
         end
        %----------------------------------------------------- extractSlide
         function obj = extractSlides(obj, varargin)
            % Find PPT files, extract slides, create new PPT(s)
            
            
            % Process input
            if nargin == 3
                % Define parameters
                obj.slidesToExtract = varargin{1};
                obj.numSlidesPerPPT = varargin{2};
            end
            
            %-------------------------------------------- Get PPT File Name
            try
                % Search pptFolderPath for PPT files
                fileListCellwLabels = ...
                    obj.GetEdfFileListInfo(obj.pptFolderPath);

                % Extract path and file name information
                filePath = fileListCellwLabels(2:end, 5);
                fileName = fileListCellwLabels(2:end, 1);
                mergePathNameF = @(x)strcat(filePath{x},fileName{x});
                numPptFiles = length(filePath);
                pptFnWithPath = arrayfun(mergePathNameF,[1:1:numPptFiles], ...
                    'UniformOutput', 0)';

                % Save Information
                obj.filePathP = filePath;
                obj.fileNameP = fileName;
                obj.numPptFilesP = numPptFiles;
                obj.pptFnWithPathP = pptFnWithPath;  
            catch
                fprintf('Could not create PPT file list\n');
                return
            end
            %--------------------------------------- Begin Slide Extraction     
            try
                % Power Point information
                numPptFiles = numPptFiles;
                slideToRemove = obj.slidesToExtract;
                numPerPPT = obj.numSlidesPerPPT;

                % Open target ppt
                fncountStart = 1;
                fncount = fncountStart;
                ppt2 = actxserver('PowerPoint.Application');
                ppt2.visible;
                fnName = sprintf('%s%s%.0f.ppt', ...
                    obj.save_folder_name, obj.pptFileNamePrefix, fncount);

                % Check if file exists
                existIndex = exist(fnName, 'file');   
                if existIndex == 2
                    % Open exisiting file
                    op2 = invoke(ppt2.Presentations,'Open',fnName); 
                else
                    % Create a new file
                    op2 = invoke(ppt2.Presentations,'Add'); 
                    ppt2.ActivePresentation.SaveAs(fnName); 
                    ppt2.ActivePresentation.Close; 
                    op2 = invoke(ppt2.Presentations,'Open',fnName); 
                end

                % Process PPT
                numPptToProcess = numPptFiles;  
                count = 0;
                
                for p = 1:numPptToProcess
                    % Next File
                    fn = obj.pptFnWithPathP{p};

                    % Open Next Presentation
                    ppt = actxserver('PowerPoint.Application');
                    ppt.visible;
                    op = invoke(ppt.Presentations,'Open', fn);       
                    My_Slides = op.Slides;

                    for s = 1:length(slideToRemove)
                        % Copy slide to combined presentation      
                        op.Slides.Item(slideToRemove(s)).Copy
                        op2.Slides.Paste;
                        op2.Save
                        count = count +1;
                        
                        % Check if we need to create new output file
                        if and(rem(count,numPerPPT) == 0,  p>fncountStart)
                            % Close Current figure
                            op2.Close

                            % Create new file name
                            fncount = fncount + 1;
                            fnName = sprintf('%s%s%.0f.ppt', ...
                                obj.save_folder_name, obj.pptFileNamePrefix, fncount);


                            % Check if file exists
                            existIndex = exist(fnName, 'file');   
                            if existIndex == 2
                                % Open exisiting file
                                op2 = invoke(ppt2.Presentations,'Open',fnName); 
                            else
                                % Create a new file
                                op2 = invoke(ppt2.Presentations,'Add'); 
                                ppt2.ActivePresentation.SaveAs(fnName); 
                                ppt2.ActivePresentation.Close; 
                                op2 = invoke(ppt2.Presentations,'Open',fnName); 
                            end
                        end                        
                        
                    end

                    % Close current presentation
                    op.Close   
                end
            catch
               fprintf('Could not access or create PPT files\n'); 
            end
            
            % Close combined presentation
            try
             op2.Close
            catch
            
            end
         end
    end
    %---------------------------------------------------- Private functions
    methods (Access=protected)
        %------------------------------------------------- Support function
        %----------------------------------------------- GetEdfFileListInfo
        function varargout = GetEdfFileListInfo(obj, varargin)
            % Create default value
            value = [];
            folderPath = '';
            xlsOut = 'edfFileList.xls';

            % Process input
            if nargin ==0
                % Open window
                folderPath = uigetdir(cd,'Set EDF search folder');    
                if folderPath == 0
                    error('User did not select folder');
                end
            elseif nargin == 2
                % Set EDF search path
                folderPath = varargin{1};
            else
                fprintf('fileStruct = obj.locateEDFs(path| )\n');
            end

            % Get File List
            fileTree  = dirr(folderPath, obj.pptSearchStr);
            [fileList fileLabels]= flattenFileTree(fileTree, folderPath);
            fileList = [fileLabels;fileList];

            % Write output to xls file
            if nargout == 0
                xlsOut = strcat(folderPath, xlsOut);
                xlswrite('edfFileList.xls',[fileLabels;fileList]);
            else
                varargout{1} = fileList;
            end

            %---------------------------------------------- FlattenFileTree
            function varargout = flattenFileTree(fileTree, folder)
                % Process recursive structure created by dirr (See MATLAB Central)
                % find directory and file entries
                dirMask = arrayfun(@(x)isstruct(fileTree(x).isdir) == 1, ...
                    [1:length(fileTree)]);
                fileMask = ~dirMask;

                % Recurse on each directory entry
                fileListD = {};
                if sum(int16(dirMask)) > 0
                   dirIndex = find(dirMask);
                   for d = dirIndex
                       folderR = strcat(folder,'\',fileTree(d).name);
                       fileListR = flattenFileTree(fileTree(d).isdir, folderR);
                       fileListD = [fileListD; fileListR];
                   end 
                end

                % Merge current and recursive list
                fileList = {};
                if sum(int16(fileMask)) > 0
                   fileIndex = find(fileMask);
                   for f = fileIndex
                       entry = {fileTree(f).name ...
                                fileTree(f).date  ...
                                fileTree(f).bytes  ...
                                fileTree(f).datenum ...
                                folder};
                       fileList = [fileList; entry];
                   end   
                end

                % Merg diretory and file list
                fileList = [fileList; fileListD];

                % Pass file list labels on export
                if nargout == 1
                    varargout{1} = fileList;
                elseif nargout == 2
                    varargout{1} = fileList;
                    varargout{2} = ...
                        {'name', 'date', 'bytes',  'datenum', 'folder'};
                end
            end
        end
    end
    %------------------------------------------------- Dependent Properties
    methods(Static)
    end   
end


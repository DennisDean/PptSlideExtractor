function testPptSlideExtractorClass
%testPptSlideExtractorClass Test function for extracting slides
%   Test function for extracting slides from a list of powerpoint files.
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
%    obj = SlideExtractorClass(pptFolderPath, pptFileNamePrefix, save_folder_name)
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
    
    

% Test Flags
test_1 = 0;       % Run with default parameters
test_2 = 0;       % Set extraction parameters to create one output file
test_3 = 1;       % Set extraction parameters to create two output file

%------------------------------------------------------------------- Test 1
if test_1 == 1
    % Echo test information to console
    test_id = 1;
    test_str = 'Run with default parameters';
    fprintf('%0.f. %s\n', test_id, test_str);
    
    % Define required parameters
    pptFolderPath = strcat(cd,'\PPT_Examples\');
    pptFileNamePrefix = 'PPT_Examples_Test_1_';
    save_folder_name = strcat(cd,'\');
    
    % Create class and extract
    secObj = PptSlideExtractorClass...
                      (pptFolderPath, pptFileNamePrefix, save_folder_name);
    secObj = secObj.extractSlides;              
end
%------------------------------------------------------------------- Test 2
if test_2 == 1
    % Echo test information to console
    test_id = 2;
    test_str = 'Set extraction parameters to create one output file';
    fprintf('%0.f. %s\n', test_id, test_str);
    
    % Define required parameters
    pptFolderPath = strcat(cd,'\PPT_Examples\');
    pptFileNamePrefix = 'PPT_Examples_Test_2_';
    save_folder_name = strcat(cd,'\');
    
    % Create class
    secObj = PptSlideExtractorClass...
                      (pptFolderPath, pptFileNamePrefix, save_folder_name);
                  
    % Set extraction parameters              
    slidesToExtract = [1, 13, 26];
    numSlidesPerPPT = 100;
       
    % Extract
    secObj = secObj.extractSlides(slidesToExtract, numSlidesPerPPT);              
end
%------------------------------------------------------------------- Test 2
if test_3 == 1
    % Echo test information to console
    test_id = 3;
    test_str = 'Set extraction parameters to create two output file';
    fprintf('%0.f. %s\n', test_id, test_str);
    
    % Define required parameters
    pptFolderPath = strcat(cd,'\PPT_Examples\');
    pptFileNamePrefix = 'PPT_Examples_Test_3_';
    save_folder_name = strcat(cd,'\');
    
    % Create class
    secObj = PptSlideExtractorClass...
                      (pptFolderPath, pptFileNamePrefix, save_folder_name);
                  
    % Set extraction parameters              
    slidesToExtract = [1, 13, 26];
    numSlidesPerPPT = 9;
       
    % Extract
    secObj = secObj.extractSlides(slidesToExtract, numSlidesPerPPT);              
end
end


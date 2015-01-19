# PptSlideExtractor
Create PowerPoint file(s) from slides extracted from PPT files identified in a folder. 

Slides can be extracted either from a script file or from a GUI. A MATLAB App (GUI), test PPT files and example output can be found in the [release](https://github.com/DennisDean/PptSlideExtractor/releases) section.

Start GUI

    >SlideExtractorFig


 Input

        pptFolderPath : Folder path that contains PPT/PPTX files
    pptFileNamePrefix : Destination PPT file name prefix
     save_folder_name : Destination Path to save destination PPT
      slidesToExtract : Array of slides to remove from each PPT
      numSlidesPerPPT : Number of slides to save in each destintation 
                        PPT. Each created PPT is numbered sequentially 
                        starting with 1.

 Function Prototypes

    obj = SlideExtractorClass(pptFolderPath, pptFileNamePrefix, save_folder_name)
    obj = obj.extractSlides
    obj = obj.extractSlides(slidesToExtract, numSlidesPerPPT)

 Requirements

    Tested with Microsoft PowerPoint 2010

 Acknowledgements

    The following open source utilities are called by
    SlideExtractorClass
   
    saveppt2
    http://www.mathworks.com/matlabcentral/fileexchange/19322-saveppt2

    dirr
    http://www.mathworks.com/matlabcentral/fileexchange/8682-dirr--find-files-recursively-filtering-name--date-or-bytes-


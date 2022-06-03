

Const ppAdvanceOnTime = 2   ' 2 = Run according to timings 
Const ppShowTypeKiosk = 3   ' 3 - Run in full screen mode 

' Create and set file system object variables for
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objCurrentPPTX = objFileSys.GetFile("add local copy path of the slide show") 'local copy of shared file
Set objNewPPTX = objFileSys.GetFile("add network path to shared show here")  'shared file


'create a powerpoint object
Set objPPTX = CreateObject("PowerPoint.Application")
objPPTX.Visible = True

'open the local slide show
Set objPresentation = objPPTX.Presentations.Open(objCurrentPPTX.Path) 'open local slide show
' Apply powerpoint settings
objPresentation.Slides.Range.SlideShowTransition.AdvanceOnTime = TRUE
objPresentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 
objPresentation.SlideShowSettings.ShowType = ppShowTypeKiosk
objPresentation.SlideShowSettings.LoopUntilStopped = True

' Run the slideshow
Set objSlideShow = objPresentation.SlideShowSettings.Run.View


'loop until error
Do Until Err <> 0  
    'compair last modify times of local and shared files 
    'if newer file on the network share copy it to the local slides shows location and restart
    
	If objNewPPTX.DateLastModified > objCurrentPPTX.DateLastModified Then
		objPresentation.Close
        objFileSys.CopyFile objNewPPTX, objCurrentPPTX, True
        
        Set objPresentation = objPPTX.Presentations.Open(objCurrentPPTX.Path)

        ' Reapply settings
        objPresentation.Slides.Range.SlideShowTransition.AdvanceOnTime = TRUE
        objPresentation.SlideShowSettings.AdvanceMode = ppAdvanceOnTime 
        objPresentation.SlideShowSettings.ShowType = ppShowTypeKiosk
        objPresentation.SlideShowSettings.LoopUntilStopped = True

      ' Rerun the slideshow
        Set objSlideShow = objPresentation.SlideShowSettings.Run.View

     End If

Loop

objPresentation.Saved = False
objPresentation.Close
objPPTX.Quit

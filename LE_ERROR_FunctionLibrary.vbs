Function ClickOKOnError()
	WpfWindow("devnamepath:=;").WpfObject("text:=Error").WpfButton("devname:=OK").highlight
End Function	
 
Function ClickOKONError(Object)
WpfWindow("devnamepath:=;").WpfObject("text:=Error").WpfButton("devname:=OK").Highlight
WpfWindow("devnamepath:=;").WpfObject("text:=Error").WpfButton("devname:=OK").Click
End Function 
 
 
Function ClickCancelOnBriefError(Object)
WpfWindow("devname:=MainWindow").WPFButton("devname:=Cancel").Click
End Function 
 

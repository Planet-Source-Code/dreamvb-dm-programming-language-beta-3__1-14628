Procedure()
	BkColour=clBlack;
	Mode(13);
	CurrentX=2000;
	CurrentY=100;
	TextColour=clYellow;
	TextOut("DM++ Beta 2");
	CurrentX=800;
	CurrentY=800;
	TextColour=clred;
	TextOut("Now Supports" & Char{10} & Char{9} & "Text Positioning...");
	Window.Show;
END.
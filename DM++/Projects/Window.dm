Procedure()
	Window.Caption="New Window";
	Delay=1;
	Mode(13);
	BkColour=clBlack;
	Window.Show;
	Beep;
	TextColour=clWhite;
	TextOut("Welcome");
	Delay=2;
	TextColour=clRed;
	Mode(12);
	TextOut("The New DM++ Beta 2.1");
	Mode(13);
	TextColour=clYellow;
	CurrentX=200;
	CurrentY=1000;
	TextOut("Now Sopports Opening Windows..");
END.

Procedure()
	BlinkText.Caption="Welcome to DM++ Beta 3";
	BlinkText(Red+Green);
	BlinkText.FontSize=24;
	BlinkText.FontName="Roman";
	BlinkText.Left=500;
	BlinkText.Top=200;
	Window.Caption="New Window";
	TextColour=clWhite;
	BkColour=clBlack;
	Mode(13);
	CurrentX=1500;
	CurrentY=2000;
	TextOut("Now Supports Blinking Text");
	Window.Show;
END.

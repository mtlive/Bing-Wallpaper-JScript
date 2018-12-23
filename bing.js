bingjs="http://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&cc=us";
hpwp="http://www.bing.com/hpwp/";

wshShell = new ActiveXObject("WScript.Shell") ;
explorer = new ActiveXObject("Scripting.FileSystemObject");
tempFolder=explorer.GetSpecialFolder(2).Path;
image=tempFolder.concat("\\bing-image.jpg");  //where to save image.
//image=wshShell.ExpandEnvironmentStrings("%appdata%\\bing-image.jpg");

objHTTP = new ActiveXObject("MSXML2.XMLHTTP");
try {
	objHTTP.Open("GET", bingjs, false);
	objHTTP.Send;
	if (objHTTP.status !== 200) 
		throw ex;
} catch (e) {	//Retry if connection failed
	WScript.Sleep(10000);
	objHTTP.Open("GET", bingjs, false);
	objHTTP.Send;
} 

var hsh = objHTTP.ResponseText.match(/"hsh":"\w*"/).toString().substr(7, 32); 
var imgurl = hpwp.concat(hsh);
var title= objHTTP.ResponseText.match(/"title":".*?"/).toString().substr(9).slice(0,-1);
var copyright= objHTTP.ResponseText.match(/"copyright":".*?"/).toString().substr(13).slice(0,-1);

//Download Image
objHTTP.Open("GET", imgurl, false);
objHTTP.Send();
if (objHTTP.status === 200) {
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Open();
  stream.Type = 1; //adTypeBinary
  
  stream.Write(objHTTP.responseBody);
  stream.Position = 0; //Set the stream position to the start
  
  if (explorer.Fileexists(image)) { explorer.DeleteFile(image); }
  
  stream.SaveToFile(image);
  stream.Close();
}


//Applying image as the wallpaper
//wshShell = new ActiveXObject("WScript.Shell") ;
wshShell.RegWrite ("HKCU\\Control Panel\\Desktop\\Wallpaper", image);
WScript.Sleep(1000);
//Refresh wallpaper
wshShell.Run("%windir%\\System32\\RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters", 1, true);
WScript.Sleep(2000);
wshShell.Run("%windir%\\System32\\RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters", 1, true);
wshShell.Run("%windir%\\System32\\RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters", 1, true);

title="Today's Wallpaper: ".concat(title);
wshShell.Popup(copyright, 0, title, 0x40);//Show image description.
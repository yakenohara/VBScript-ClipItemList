' <# License>------------------------------------------------------------
' 
'  Copyright (c) 2018 Shinnosuke Yakenohara
' 
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
' 
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
' 
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>
' 
' -----------------------------------------------------------</License #>

'設定
styleStr_DirStart = """"
styleStr_DirEnd = """"
styleStr_MiddleOfList = "┣"
styleStr_EndOfList = "┗"

'ライブラリからオブジェクト生成
Set ShellObj = CreateObject("WScript.Shell")
Set FSObj = createObject("Scripting.FileSystemObject")
Set folderArrayList = CreateObject("System.Collections.ArrayList")
Set fileArrayList = CreateObject("System.Collections.ArrayList")

parentDirStr = ""

For Each arg In WScript.Arguments
    
    'ディレクトリ名を保存
    If parentDirStr <> "" Then '2回目以降のディレクトリの場合

        '複数のディレクトリを指定していないかどうかチェック
        If parentDirStr <> FSObj.getParentFolderName(arg) Then
            WScript.Echo "複数ディレクトリに渡るパラメータ指定はできません"
            WScript.Quit '中断
        End If

    Else '一回目のループの場合はディレクトリ名を保存
        parentDirStr = FSObj.getParentFolderName(arg)

    End If

    If FSObj.FolderExists(arg) Then 'フォルダの場合
        Call folderArrayList.Add(FSObj.GetFileName(arg))

    Else 'ファイルの場合
        Call fileArrayList.Add(FSObj.GetFileName(arg))

    End If
    
Next

'名前順で並べ替え
Call folderArrayList.Sort
Call fileArrayList.Sort

'結合して配列を取得
folderArrayList.AddRange(fileArrayList)
names = folderArrayList.ToArray

'ItemListの作成
printStr = styleStr_DirStart & parentDirStr & styleStr_DirEnd
itrMx = UBound(names)
For itr = 0 To (itrMx - 1)
    printStr = printStr & vbCrLf & styleStr_MiddleOfList & names(itr)
Next    
printStr = printStr & vbCrLf & styleStr_EndOfList & names(itr)

'クリップボードにコピー
ShellObj.Exec("clip").StdIn.Write printStr

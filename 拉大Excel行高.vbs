

'编码目的
'打印Excel表格时，经常出现编辑时显示的页面和打印预览显示的的不一致，而打印预览就是最终在纸上显示的内容
'通常表现为，打印预览的表格，某些单元格最上面或最下面只能显示字的一半，甚至最后面会少字，但编辑状态下没有问题
'概括来讲，Excel所见非所得

'这个困扰，时不时地出现在打印招投标文件、产品技术参数说明书等场景下。要命的是，这些打印场景，不允许半点差错
'即便你设置了自动适应行宽、自动适应行高，也不一定能解决问题，字被“吃掉”的现象依然存在
'表格数据较少的时候，也许可以通过人工逐行检查，手工拉高行高能够解决（一般不会拉宽，因为你不会希望内容跨页）
'但如果数据非常大，人眼比对，不仅效率低，效果也不好

'Excel编辑界面和打印预览不一致，本质原因是显示器和打印机的DPI不对应
'Excel表格的宽度，以字符和像素数表示，高度以磅和像素数表示，并不是物理上的长度单位
'运气好，显示器和打印机DPI适配，所见即所得；运气不好，部分字就会被“吃掉”
'目前，可以通过调正显示器分辨率（一般是降低），比如，23.8寸 1920*1080 的显示器，win7及以上操作系统，使用 Adobe PDF 打印机标准模式时，设置显示字体放大125%，屏幕显示的和打印的就一致起来
'但调正分辨率，不仅麻烦、影响显示效果，而且成功率低

' 此程序不能完全保证打印出来的不吞字，因为你的列宽太窄、文字太多，到时候会涉及跨页问题
' 经测试发现，同样的内容，同样的单元格宽高，在不同的工作表中，打印预览的效果还不一样。有的不吞字，有的吞字，这就无解了
' 所以本程序适用于单元格高度不太高、不存在跨页问题的场景。打印大文字表格，真的用Word!

' Excel真的不适合展示大量文本。写技术参数这些大文字量的东西，还是用Word吧！
'
' -------------------   以上编码目的没实现，能够实现的功能看下面  ---------------------------------
'


' 本程序能够实现：
' 在用户指定的列宽下，批量适度增加行高，确保在 Excel 编辑界面，文字内容不会顶着单元格边界，至少会留有一定的空白，看起来相对美观
' 在工作表行数多、但单个单元格文字量不大的场景下，杜绝编辑界面无误、但打印出来的字被“吞掉”的现象

' 本程序不能实现：
' 单元格填充大量文字的场景下，打印出来的页面能完整显示单元格的所有内容（Excel 打印跨页问题）

' 本程序特色：
' 具体该拉多高，由该行字号大小、该行单元格文字量确定(不是写死的固定值)。字号越大，文字量越多，拉出来的空白越多
' 必须要缩小字号才能拉高的单元行，会用黄底色标出
' 即使缩小了字号，也无法按照设定好的比例拉高的单元行，会用红底色标出
'


' 整体思路，根据单元格字号大小，保证单元格内容的上下加起来有足够高度的空白。需要：
' 遍历Worksheet 的所有行
' 获取单元格的字号（遍历单元格）
' 针对某一行，所有单元格中，最大/小字号
' 获取某一行当前的行高
' 以当前行高和单元格字号为基准,设置新的行高


Sub 增加行高()

    Call AddHeight(Application.ActiveSheet)

End Sub

'
'调整单个工作表中的所有行高
' @param { Worksheet } myWorksheet 需要调整行高的工作表
' @return 无
'

Function AddHeight(ByRef myWorksheet As Worksheet)

    Application.ScreenUpdating = False

    Dim myRange As Range
    Set myRange = myWorksheet.UsedRange
    
'    工作簿中的行数和列数
    Dim cRows As Integer, cColumns As Integer
    cRows = myRange.Rows.Count
    cColumns = myRange.Columns.Count
    
'    警告信息，提示用户有哪些行被缩小了字体才调节成功的
    Dim warningMsg As String
    warningMsg = ""
     
'    失败信息，告知用户有哪些行即便缩小了字体，调节还是失败的
    Dim errorMsg As String
    errorMsg = ""
    
'    整合警告信息和失败信息
    Dim popMsg As String
    popMsg = ""

'    从第1行开始，逐行调整行高
    For i = 1 To cRows Step 1
        Call AddSingleRowHeight(myRange.Rows(i), i, cColumns, warningMsg, errorMsg)
    Next
    
    If Len(warningMsg) > 0 Then
        popMsg = popMsg + "第 " + warningMsg + "行调小了字号才拉大了行高，已用黄底色标出，请知悉" + vbLf
    End If
    
    If Len(errorMsg) > 0 Then
        popMsg = popMsg + "第 " + errorMsg + "字号已经调成最小了还是无法拉大行高，已用红底色标出，请知悉" + vbLf + vbLf
    End If
    
    Application.ScreenUpdating = True
    
    If Len(popMsg) > 0 Then
        popMsg = popMsg + "以上 黄/红 底色标出的行，请人工检查！"
        MsgBox (popMsg)
    End If
    
End Function

'
'调整单个行高
' @param { Range} row 需要调整高度的行
' @param { Integer } i 需要调整高度的行号
' @param { Integer } cColumns 列的数量
' @param { String } warningMsg 警告信息
' @param { String } errorMsg 失败信息
' @return 无
'

Function AddSingleRowHeight(ByRef row As Range, ByVal i As Integer, ByVal cColumns As Integer, ByRef warningMsg As String, ByRef errorMsg As String)


'   先自适应行高
    row.AutoFit
    
'   自适应后的行高
    Dim originH As Single
    originH = row.RowHeight
   
'   存放某一行中，每个单元格的字体值
    Dim hArray() As Single
    ReDim hArray(cColumns - 1)
    
'   最大字体值
    Dim maxSize As Single

'   遍历该行，取该行单元格的最大字体值
    Dim j As Integer
    j = 0
    For Each ran In row.Cells
        hArray(j) = ran.Font.Size
        j = j + 1
    Next
    maxSize = GetMax(hArray)
    
'   确定调整后的行高
    Dim newH As Single
    newH = originH + maxSize * GetTimes(originH)
    
'   行高有最大值409!超过该值无法设定
'   自适应行高后，发现行高超过最大限值，尝试把该行字体缩小1号，再重新自适应、拉高行
'   以该行最小的字号为基准，缩小1号
        
    If newH > 409 Then
    
        Do While newH > 409
            Dim minSize As Single, newSize As Single
            
    '        遍历行，取该行单元格的最小值
            j = 0
            For Each ran In row.Cells
                hArray(j) = ran.Font.Size
                j = j + 1
            Next
            minSize = GetMin(hArray)
            
            newSize = minSize - 1
            
    '       缩小后的字号，最小为6，不能再小了
            If newSize >= 6 Then
                
    '            把该行所有单元格设置成新的size
                For Each ran In row.Cells
                    ran.Font.Size = newSize
                Next
                
                row.AutoFit
                newH = row.RowHeight + newSize * GetTimes(row.RowHeight)
            Else
                Exit Do
            End If
            
        Loop
         
        If newH > 409 Then
'            调整过的、且调整失败的行，背景色设置红色
            row.Interior.Color = 255
            errorMsg = errorMsg + CStr(i) + "  "
        Else
'            调整过的、且调整成功的行，背景色设置黄色
            row.Interior.Color = 65535
            warningMsg = warningMsg + CStr(i) + "  "
        End If
    
    End If
    
'    调整后的行高如果不超过最大值，执行拉高行高操作
    If newH <= 409 Then
        row.RowHeight = newH
    End If
            
End Function


'获取数组中的最大值

Function GetMax(arr() As Single) As Single

    Dim temp As Single
    temp = arr(0)
    For Each x In arr
        If x > temp Then
            temp = x
        End If
    Next
    
    GetMax = temp
End Function

'获取数组中的最小值

Function GetMin(arr() As Single) As Single

    Dim temp As Single
    temp = arr(0)
    For Each x In arr
        If x < temp Then
            temp = x
        End If
    Next
    
    GetMin = temp
End Function

' 根据原行高，确定增加的行高是字号的多少倍
' 这里可以修改文本在单元格内的留白大小

Function GetTimes(h As Single) As Integer
    Dim num As Single
    num = h / 50
    
    If num > 0 And num <= 1 Then
        GetTimes = 1
    ElseIf num > 1 And num <= 3 Then
        GetTimes = 2
    ElseIf num > 3 And num <= 6 Then
        GetTimes = 3
    ElseIf num > 6 And num <= 8 Then
        GetTimes = 4
    ElseIf num > 8 Then
        GetTimes = 5
    End If
   
End Function












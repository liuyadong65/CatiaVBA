
<img width="790" height="937" alt="image" src="https://github.com/user-attachments/assets/32d3d201-8052-480a-91ed-c6e31458ba23" />

帮助文档：
C:\Users\YLIU119\OneDrive - azureford\desktop\Catia tool\online

## 创建catiavba 文档
<img width="640" height="300" alt="image" src="https://github.com/user-attachments/assets/241d6c86-b256-40ad-9c1b-a0b2dffaf52d" />


### 第一个测试文件，获取文件名：

    Sub GetCatiaChild()
    
    Dim oproductDoc As Document
    Set oproductDoc = CATIA.ActiveDocument ' returns a Document
    
    Set oProduct = oproductDoc.Product
    
    Dim TopProductName As String
    TopProductPN = oProduct.PartNumber
    
    MsgBox TopProductPN
    
    If oProduct.Products.Count >= 1 Then
        Set oSubProduct = oProduct.Products.Item(1)
        MsgBox oSubProduct.PartNumber
    End If
    
    End Sub


### 递归遍历 所有文件名
    
Option Explicit

' === 入口过程：直接运行它 ===
Sub ExportProductTreeToText()

    ' 确保当前是装配文档（CATProduct）
    If CATIA Is Nothing Then
        MsgBox "未检测到 CATIA 应用。", vbExclamation
        Exit Sub
    End If
    
    If TypeName(CATIA.ActiveDocument) <> "ProductDocument" Then
        MsgBox "请先打开一个装配文档（.CATProduct）。当前文档类型：" & TypeName(CATIA.ActiveDocument), vbExclamation
        Exit Sub
    End If

    Dim prodDoc As ProductDocument
    Set prodDoc = CATIA.ActiveDocument
    
    Dim rootProd As Product
    Set rootProd = prodDoc.Product
    
    ' 收集每一行文本
    Dim lines As Collection
    Set lines = New Collection
    
    ' 递归遍历
    WalkProductTree rootProd, 0, lines
    
    ' 合并为文本
    Dim outText As String
    outText = Join(CollectionToArray(lines), vbCrLf)
    
    ' 保存到桌面
    Dim outPath As String
    outPath = "C:\Users\YLIU119\OneDrive - azureford\vscode\catiaVBA\ProductTree.txt"
    
    Dim f As Integer
    f = FreeFile
    Open outPath For Output As #f
    Print #f, outText
    Close #f
    
    MsgBox "已导出父子结构到：" & outPath, vbInformation

End Sub

' === 递归遍历 ===
' prod  : 当前产品节点
' level : 层级深度（用于缩进）
' lines : 收集输出行
Private Sub WalkProductTree(ByVal prod As Product, ByVal level As Long, ByRef lines As Collection)
    Dim indent As String
    indent = String$(level * 2, " ")    ' 每层两个空格缩进（可调）
    
    Dim hasChildren As Boolean
    hasChildren = (prod.Products.Count > 0)
    
    ' 取显示名：优先 PartNumber，若为空则用 Name
    Dim displayName As String
    displayName = SafePartNumber(prod)
    If Len(Trim$(displayName)) = 0 Then displayName = prod.Name
    
    ' 有子零件前加 "+ "
    Dim line As String
    line = indent & IIf(hasChildren, "+ ", "  ") & displayName
    lines.Add line
    
    ' 如果有子零件，继续递归
    If hasChildren Then
        Dim i As Long
        Dim child As Product
        For i = 1 To prod.Products.Count          ' 注意：CATIA 集合是 1 基
            Set child = prod.Products.Item(i)
            WalkProductTree child, level + 1, lines
        Next i
    End If
End Sub

' === 安全获取 PartNumber（避免异常） ===
Private Function SafePartNumber(ByVal prod As Product) As String
    On Error Resume Next
    SafePartNumber = prod.PartNumber
    On Error GoTo 0
End Function

' === 将 Collection 转为字符串数组，便于 Join ===
Private Function CollectionToArray(ByVal col As Collection) As String()
    Dim arr() As String
    Dim i As Long
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col.Item(i)
    Next i
    CollectionToArray = arr
End Function

## 获取Part
    Sub GetCatiaChild()
    
    
    Dim oproductDoc As Document
    Set oproductDoc = CATIA.ActiveDocument ' returns a Document
    
    Set oProduct = oproductDoc.Product
    
    Dim TopProductName As String
    TopProductPN = oProduct.PartNumber
    
    
 
    
    If oProduct.Products.Count >= 1 Then
        Set oSubProduct = oProduct.Products.Item(1)
        
        
        Set leafProd = oSubProduct.Products.Item(1) ' 这是你遍历到的叶子 Product ,到一个没有子项的 Product（通常是零件）
        
        
       ' Dim leafProd As Product        ' 这是你遍历到的叶子 Product

        If leafProd.Products.Count = 0 Then
            Dim refProd As Product
            Set refProd = leafProd.ReferenceProduct
    
            If Not refProd Is Nothing Then
                If TypeName(refProd.Parent) = "PartDocument" Then
                    Dim partDoc As PartDocument
                    Set partDoc = refProd.Parent    'Product 的父链拿到它所属的文档对象 零件文档 —— PartDocument,装配文档 —— ProductDocument
            
                    ' 现在可以拿 Part / Bodies 等
                    Dim partObj As Part
                    Set partObj = partDoc.Part
            
                    Dim bodies As bodies
                    Set bodies = partObj.bodies
                
                End If
             End If
        End If
    End If
    
    End Sub



    ' If TypeName(CATIA.ActiveDocument) = "PartDocument" Then
    ' Dim partDoc As PartDocument
    ' Set partDoc = CATIA.ActiveDocument        ' ← 已经有实例
    ' ' 此时就不需要 Set partDoc = refProd.Parent 了
    ' End If


## 提取body的面
现在的做法是对每个 Body 先把所有 Face 搜出来，再逐个建 Extract，这在几何复杂/面数多的时候会非常慢。
其实你要的“从第一选中的面开始，按点联系（Point Continuity）把整片面一并提取”可以直接用 HybridShapeExtract 的**传播（Propagation）**功能一次完成，不必循环每个面。

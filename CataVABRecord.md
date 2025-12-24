
<img width="790" height="937" alt="image" src="https://github.com/user-attachments/assets/32d3d201-8052-480a-91ed-c6e31458ba23" />

帮助文档：
C:\Users\YLIU119\OneDrive - azureford\desktop\Catia tool\online

## 创建catiavba 文档
<img width="640" height="300" alt="image" src="https://github.com/user-attachments/assets/241d6c86-b256-40ad-9c1b-a0b2dffaf52d" />


### 选路径打开文档：
    Sub CATMain()
    
    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
    
    Dim productDocument1 As ProductDocument
    Set productDocument1 = documents1.Open("C:\Users\YLIU119\OneDrive - azureford\desktop\SeatCAD\R2TB-R61178-A.CATProduct")
    
    Dim product1 As Product
    Set product1 = productDocument1.Product
    
    product1.ApplyWorkMode DESIGN_MODE
    
    Dim specsAndGeomWindow1 As SpecsAndGeomWindow
    Set specsAndGeomWindow1 = CATIA.ActiveWindow
    
    Dim viewer3D1 As Viewer3D
    Set viewer3D1 = specsAndGeomWindow1.ActiveViewer
    
    viewer3D1.Reframe
    
    Dim viewpoint3D1 As Viewpoint3D
    Set viewpoint3D1 = viewer3D1.Viewpoint3D
    
    End Sub



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


## 无参化数据：
在 CATIA V5 里，把由 HybridShapeExtract 得到的曲面无参化，常用两种做法：

创建 Datum（显式曲面）：用 HybridShapeFactory.AddNewSurfaceDatum 把 Extract 结果转换为显式（Explicit）曲面，从而断开与原几何的设计规格关联；该显式对象的类型是 HybridShapeSurfaceExplicit（文档明确指出它由 AddNewSurfaceDatum 创建）。 [catiadesign.org], [r1 HybridS...ct) - Free]
Paste Special - As Result：把几何复制后以“结果（无链接）”方式粘贴，得到独立的几何实体（宏里可用 CATPrtResultWithOutLink 标识）。这同样能无参化，但更适合跨 Part／跨容器复制的场景


## 平面的坑
你观察到 TypeName(sel.Item2(t).Value) 有时是 PlanarFace，而且这个“面”其实来自坐标系的基准平面（XY/YZ/ZX）而不是实体/曲面上的真实拓扑面——这是一个常见坑。
为什么会这样？

Topology.Face/Topology.PlanarFace 是拓扑层面的“面单元”概念，既可能来自实体/曲面（真·BRep 面），也可能来自显式的基准平面/辅助几何（比如坐标系里的平面），在某些数据结构或搜索范围的组合下，搜索确实会返回“PlanarFace”对象。
仅用 TypeName = "PlanarFace" 并不能保证它就是“可用于 HybridShapeExtract 的真实面”（你已经遇到了：继续执行到 Update 时才失败）。


解决思路：在“取种子面”那一步就剔除基准平面，保留真实 BRep 面
可以在尝试 AddNewExtract 之前，对候选对象做两层校验：

父对象校验：该面必须属于当前 body（或者属于一个 Shape/Surface），不能来自 AxisSystem、HybridShapePlaneExplicit 这类基准平面。
可测量性校验：对该 Reference 做一次轻量测量（面积），真实的 BRep 面通常可以正常返回面积；若抛错或面积为 0，则认为不是可用面。


注：这两步在宏层面非常轻量，不会像到 part.Update 才报错那样浪费时间。

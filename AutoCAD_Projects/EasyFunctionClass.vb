Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports Newtonsoft
Imports System.Drawing
Imports System.IO

Public Class EasyFunctionClass
    <CommandMethod("HelloWorld")>
    Public Sub HelloWorld()
        'Test the netload
        'after load the AutoCAD_Projects.dll file, the "HelloWorld" should become a command in AutoCAD
        'and you can type the command and on the Editor page it will show "HelloWorld" msg

        Dim ac_doc As Document = Application.DocumentManager.MdiActiveDocument
        ac_doc.Editor.WriteMessage("HelloWorld!")

    End Sub
    <CommandMethod("PositionSizeAppWindow")>
    Public Sub PositionSizeAppWindow()
        'position the AutoCAD app main windows to a location points
        'size the AutoCAD app main windows by the windows size

        Dim point_app As Windows.Point = New Windows.Point(0, 0)
        Core.Application.MainWindow.DeviceIndependentLocation = point_app

        Dim size_app As Windows.Size = New Windows.Size(400, 400)
        Core.Application.MainWindow.DeviceIndependentSize = size_app

    End Sub
    <CommandMethod("MinimizeWindow")>
    Public Sub MinimizeWindow()
        'minimize the app window
        Core.Application.MainWindow.WindowState = Window.State.Minimized
        MsgBox("Minimized", MsgBoxStyle.SystemModal, "MinMax")
    End Sub
    <CommandMethod("WindowsState")>
    Public Sub WindowsState()
        'show autocad windows state in a msg box
        System.Windows.Forms.MessageBox.Show(Application.MainWindow.WindowState.ToString, "windows state")
    End Sub
    <CommandMethod("SizeDocWindow")>
    Public Sub SizeDocWindow()
        'change CAD drawing window
        Dim active_doc As Document = Application.DocumentManager.MdiActiveDocument
        active_doc.Window.DeviceIndependentLocation = New System.Windows.Point(3, 3)
        active_doc.Window.DeviceIndependentSize = New System.Windows.Size(500, 500)

        'Dim ac_doc_manager As DocumentCollection = Application.DocumentManager
        'Dim ac_doc As Document = DocumentCollectionExtension.Add(ac_doc_manager, "acad.dwt")
        'ac_doc_manager.MdiActiveDocument = ac_doc

    End Sub
End Class

Public Class DocOperation
    <CommandMethod("CreateDrawing")>
    Public Sub CreateDrawing()
        Dim ac_doc_mananger As DocumentCollection = Application.DocumentManager
        Dim ac_doc As Document = DocumentCollectionExtension.Add(ac_doc_mananger, "drawing_a.dwg")
    End Sub

End Class

Public Class DBOperation
    <CommandMethod("ShowBlkTbl")>
    Public Sub ShowBlkTlb()
        Dim ac_doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ac_db As Database = ac_doc.Database
        ac_doc.Editor.WriteMessage(ac_doc.Name)
        Using ac_trans As Transaction = ac_db.TransactionManager.StartTransaction()

            'get all the blocktable records from database block table
            Dim btr_list As List(Of BtRecord) = GetBtRecordList(ac_trans, ac_db)
            Dim btrlist_json As String = Json.JsonConvert.SerializeObject(btr_list, Json.Formatting.Indented)
            File.WriteAllText("c:\code_output\btrlist.txt", btrlist_json)

            'get all objects from the blocktable record *Model_Space
            Dim btr_object_list As List(Of BtrObject) = GetBtrObjectsIDLists(ac_trans, ac_db, "*Model_Space")
            Dim btrobjectlist_json As String = Json.JsonConvert.SerializeObject(btr_object_list, Json.Formatting.Indented)
            File.WriteAllText("c:\code_output\btrobjectlist.txt", btrobjectlist_json)

            Dim blk_ref_list As List(Of BtrObject) = GetBlockReferenceObjectList(btr_object_list)
            Dim btr_reflist_json As String = Json.JsonConvert.SerializeObject(blk_ref_list, Json.Formatting.Indented)
            File.WriteAllText("c:\code_output\BlkRefList.txt", btr_reflist_json)
        End Using
    End Sub

    Public Function GetBtRecordList(ByRef ac_trans As Transaction, ByRef ac_db As Database) As List(Of BtRecord)
        Dim btr_list As List(Of BtRecord) = New List(Of BtRecord)
        Dim blk_table As BlockTable = ac_trans.GetObject(ac_db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = New BlockTableRecord
        For Each btr_id As ObjectId In blk_table
            Dim btrdata As BtRecord = New BtRecord
            btrdata.btr_id = btr_id
            btr = ac_trans.GetObject(btr_id, OpenMode.ForRead)
            btrdata.btr_name = btr.Name
            btrdata.btr_class = btr_id.ObjectClass.Name
            btr_list.Add(btrdata)
        Next
        Return btr_list
    End Function
    Public Function GetBtrObjectsIDLists(ByRef ac_trans As Transaction, ByRef ac_db As Database, ByRef btr_name As String) As List(Of BtrObject)
        Dim object_list As List(Of BtrObject) = New List(Of BtrObject)
        Dim blk_table As BlockTable = ac_trans.GetObject(ac_db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = ac_trans.GetObject(blk_table(btr_name), OpenMode.ForRead)
        For Each object_id In btr
            Dim btr_object As BtrObject = New BtrObject
            btr_object.object_id = object_id
            btr_object.object_class = object_id.ObjectClass.Name
            object_list.Add(btr_object)
        Next

        Return object_list
    End Function
    Public Function GetBlockReferenceObjectList(ByRef btr_object_list As List(Of BtrObject)) As List(Of BtrObject)
        Dim blk_ref_object_list As New List(Of BtrObject)
        For Each btr_object As BtrObject In btr_object_list
            If (btr_object.object_class = "AcDbBlockReference") Then
                blk_ref_object_list.Add(btr_object)
            End If
        Next
        Return blk_ref_object_list
    End Function
End Class

Public Class BtRecord
    Public btr_id As ObjectId
    Public btr_name As String
    Public btr_class As String

End Class
Public Class BtrObject
    Public object_id As ObjectId
    Public object_class As String
End Class
Public Class AcBlockRefInfo

End Class

Public Class PawsBlockInfo

End Class

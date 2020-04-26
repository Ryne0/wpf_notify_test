Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Office.Interop
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Class MainWindow

    Dim wsl As WindowState

    Dim gFolderName As String = "[程式]傳票用資料夾"

    Dim gOlApp As Outlook.Application
    Dim gObC_DATA As ObservableCollection(Of QuerySQL.RowCase)
    Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。

        wsl = WindowState

        mIcon()

    End Sub

    Private Sub mIcon()


        'Dim nIcon As NotifyIcon = New NotifyIcon

        'nIcon.BalloonTipText = "Hello, 檔案監視器" '/設定程式啟動時顯示的文字
        'nIcon.Text = "檔案監視器" '最小化到托盤時，滑鼠點選時顯示的文字
        'nIcon.Icon = New System.Drawing.Icon("test.png") '// 程式圖示
        'nIcon.Visible = True


        'AddHandler nIcon.MouseDoubleClick, AddressOf OnNotifyIconDoubleClick

        'nIcon.ShowBalloonTip(1000)

    End Sub

    Private Sub OnNotifyIconDoubleClick(sender As Object, e As MouseEventArgs)
        Show()
        WindowState = wsl
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        'Dim s = "e:\letter.ico"

        'Me.WindowState = System.Windows.WindowState.Minimized
        'Me.nIcon.Icon = New Icon(s)
        '' Me.nIcon.Icon = Environment.CurrentDirectory
        'Me.nIcon.ShowBalloonTip(5000, "Hi", "This is a BallonTip from Windows Notification", ToolTipIcon.Info)

    End Sub

    Private Sub Window_StateChanged(sender As Object, e As EventArgs)
        wsl = WindowState
        If wsl = WindowState.Minimized Then

            Hide()
        End If
    End Sub

    Private Sub OnImageButtonClick()

    End Sub


    Enum MailEventFlag
        INI_MAIL
        NEW_MAIL
    End Enum

    '
    '  程式開啟時 初始 讀取 信件內容
    '
    Private Sub Init_Mail_Event()

        gOlApp = New Outlook.Application

        '初始化  "[程式]傳票資料夾
        IniMailFolder()

        '讀取並複製郵件到程式用收件夾 + 檢查信件夾重複
        CopyMailToFolder()

        '更新設定資料夾
        UpdateProgInbox(MailEventFlag.INI_MAIL)


        '收信事件
        AddHandler gOlApp.NewMail, AddressOf NewMailEvent


    End Sub

    '
    '  更新
    ' 讀取 [程式]傳票資料夾 更新項目
    '
    Private Sub UpdateProgInbox(ByVal Mail_Event As MailEventFlag)

        Try

            Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")
            '指向收件夾
            Dim mInbox As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            '指向收件夾 "[程式]傳票資料夾"
            Dim mProgFolder As Outlook.Folder = mInbox.Folders.Item(gFolderName)
            'mail信件 Items
            Dim mProgItems As Outlook.Items = mProgFolder.Items

            ' /*-------------- 讀取 [程式]傳票資料夾 信件  + 刪除過久郵件 迴圈  ---------------
            ' Dim ls_StoreRead As New ArrayList

            Dim dic_StoreRead As New Dictionary(Of String, QuerySQL.RowCase)
            For i = mProgItems.Count To 1 Step -1

                If i = 1 Then
                    Console.WriteLine()
                End If


                Dim mMailItem = mProgItems.Item(i)

                Dim org_Title As String = mMailItem.Subject

                mMailItem.Subject = "CheckFlag--" & org_Title
                mMailItem.Save()


                '轉換信件資料
                Dim mItemData As QuerySQL.RowCase = ParseMail(mProgItems.Item(i))

                If Not CheckDeleteMail(mItemData, mProgItems.Item(i)) Then

                    Dim MailReadStatus As Boolean = mProgItems.Item(i).UnRead

                    '信件狀態賦值
                    mItemData.Mail_UnRead_State = MailReadStatus


                    '檢查是否已有資料 
                    If dic_StoreRead.ContainsKey(mItemData.EIP_CODE_WIP) Then

                        '已有儲存相同單號  , 比對接收日期 , 較晚者取代

                        Dim mOrgRowCase As QuerySQL.RowCase = dic_StoreRead.Item(mItemData.EIP_CODE_WIP)

                        '已儲存的信件接收時間
                        Dim mOrg_DT As DateTime = mOrgRowCase.MailRecTime

                        '目前讀取時間
                        Dim mCurr_DT As DateTime = mItemData.MailRecTime

                        If DateTime.Compare(mCurr_DT, mOrg_DT) >= 0 Then  '目前讀取時間比儲存時間晚    
                            dic_StoreRead.Item(mItemData.EIP_CODE_WIP) = mItemData

                            '刪除重複單號且日期較早信件
                            TargetDeleteSingleMail(mOrgRowCase)
                        Else

                            '不賦值但仍需刪除重複信
                            TargetDeleteSingleMail(mItemData)

                        End If

                    Else
                        dic_StoreRead.Add(mItemData.EIP_CODE_WIP, mItemData)
                    End If

                    ' ls_StoreRead.Add(mItemData)
                    'gObC_DATA.Add(mItemData)
                End If

                '恢復原標題
                mMailItem.Subject = org_Title
                mMailItem.Save()


            Next
            '--------------- 讀取 [程式]傳票資料夾 信件+ 刪除過久郵件迴圈 ---------------*/

            '執行刪除標註信件
            ExcuteDeleteMail()


            '
            ' 查詢IDB   - 因應全昆傳票起始或接收需要不顯示 , 需一開始得知是否為全昆    
            ' 將查詢IDB function移至此處
            '           
            '/*---------------  合併查詢品目字串 ----------------------
            Dim query_string As String = String.Empty
            Dim query_string_tp As String = String.Empty   ' TP 件使用


            Dim arls_StoreRead As New ArrayList
            For Each Item As QuerySQL.RowCase In dic_StoreRead.Values

                If Item.Mail_Title.Contains("TP測試") Then
                    query_string_tp &= "*'" & Item.EIP_CODE_WIP & "'*"
                Else
                    query_string &= "*'" & Item.EIP_CODE_WIP & "'*"
                End If

                arls_StoreRead.Add(Item)

            Next

            'If Not Item.Flag_Add Then
            '    Continue For
            'End If

            Console.WriteLine()
            '-----------------  合併查詢品目字串 ----------------------*/

            ''查詢非TP件
            'Dim arls_result As ArrayList = gSQL_Func.QueryIDB_Muti(arls_StoreRead, query_string)
            'Dim arls_result As ArrayList = GenTestSample()
            ''查詢 TP件
            'arls_result = gSQL_Func.QueryIDB_TP(arls_result, query_string_tp)


            '排序最後結過
            'SortDateTime(arls_result)


            ' 同傳票可能修改數量發送
            '2020/04/23  修改有重複則覆蓋原資料
            '
            '/*------------- 檢查 原UI內資料 , 有重複則設定不加入 --------------         

            'For Each tmp_storeData As QuerySQL.RowCase In ls_StoreRead

            '    For Each gItem As QuerySQL.RowCase In gObC_DATA

            '        'If gItem.EIP_CODE_WIP.Equals(tmp_storeData.EIP_CODE_WIP) And
            '        '    gItem.RegTime.Equals(tmp_storeData.RegTime) And
            '        '    gItem.MailRecTime.Equals(tmp_storeData.MailRecTime) Then

            '        If gItem.EIP_CODE_WIP.Equals(tmp_storeData.EIP_CODE_WIP) And
            '           gItem.RegTime.Equals(tmp_storeData.RegTime) Then

            '            Dim org_DT As DateTime = gItem.MailRecTime       ' 已在列表中Item
            '            Dim Inbox_DT As DateTime = tmp_storeData.MailRecTime   ' 傳票資料夾  Item

            '            '覆蓋原本資料
            '            If org_DT.CompareTo(Inbox_DT) > 0 Then

            '                gItem = tmp_storeData
            '                tmp_storeData.Flag_Add = False        '檢查是否加入list 旗標      

            '            End If


            '            Exit For


            '        End If

            '    Next
            'Next


            'Dim tmp_Arls_AddData As New ArrayList   '放置無重複即將加入列表資料
            'For Each data As QuerySQL.RowCase In ls_StoreRead

            '    If data.Flag_Add Then
            '        tmp_Arls_AddData.Add(data)
            '    End If

            'Next

            '--------------- 檢查 原UI內資料 , 有重複則不加入 --------------*/

            'If arls_result.Count > 0 Then
            '    '跨執行緒更新UI 
            '    Try
            '        Dispatcher.BeginInvoke(New Action(Of ArrayList, MailEventFlag)(AddressOf Mail_Update_UI), arls_result, Mail_Event)

            '    Catch ex As Exception

            '    End Try

            'End If

        Catch ex As Exception

        End Try


    End Sub



    '
    ' 執行 刪除資料夾 刪除已標記 "Delete Mail"   
    '
    Private Sub ExcuteDeleteMail()
        Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")
        Dim mDeleteMail As Outlook.MailItem = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems).Items.Find("[Subject] = Delete Mail")
        mDeleteMail.Delete()
    End Sub

    '
    ' 用於  UpdateProgInbox 標記單號重複且日期較舊信件
    '
    Private Sub TargetDeleteSingleMail(mSenderData As QuerySQL.RowCase)


        Try
            Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")

            '指向刪除收件夾
            Dim mDeleteFolder As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)

            Dim mInbox As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)  '指向收件夾

            Dim ProgFolder As Outlook.Folder = mInbox.Folders.Item(gFolderName)  '指向 [程式] 傳票收件夾
            Dim ProgItems As Outlook.Items = ProgFolder.Items 'mail信件 Items     


            '帶入信件接收時間 
            Dim mCmpDate As String = mSenderData.MailRecTime
            'Dim mCmpDate As String = mData.MailRecTime.ToString  


            '/*------------------ 讀取主收件夾迴圈 -------------------------
            For Each mMailItem In ProgItems

                '偵測到特定標題跳出
                Dim chk_title As String = mMailItem.Subject

                If chk_title.Contains("CheckFlag--") Then
                    Continue For
                End If

                '非信件跳出
                If Not TypeOf mMailItem Is Outlook.MailItem Then
                    Continue For
                End If

                mMailItem = CType(mMailItem, Outlook.MailItem) '轉型MailItem 


                '轉換信件資料
                Dim mItemData As QuerySQL.RowCase = ParseMail(mMailItem)

                '主旨
                Dim mTitle As String = mItemData.Mail_Title
                ' Dim mTitle As String = mItemData.su

                '程式信件接收時間
                Dim mProgDate As String = mItemData.MailRecTime
                ' Dim mProgDate As Date = mMailItem.ReceivedTime
                ' Dim mProgDate As String = mMailItem.ReceivedTime.ToString   

                'If DateTime.Compare(mCmpDate, mProgDate) = 0 And mTitle.Contains(mData.Mail_Title) Then
                If mCmpDate.Equals(mProgDate) And mTitle.Contains(mSenderData.Mail_Title) Then

                    '標記  Delete Mail 移至刪除資料夾
                    mMailItem.Subject = "Delete Mail"
                    mMailItem.Save()
                    'mMailItem.Move(mDeleteFolder)


                    Exit For

                End If


            Next
        Catch ex As Exception

        End Try



    End Sub

    '
    ' 新增Mail 跨執行緒更新UI
    '         
    'Private Sub Mail_Update_UI(ByVal mData As QuerySQL.RowCase, ByVal mFlag As MailEventFlag)
    '  輸入 :　mData 預定加入列表暫存　 , 　mFlag 呼叫來源功能   ini_mail(初始執行)    new_mail  (取得Mail執行)
    Private Sub Mail_Update_UI(ByVal arls_result As ArrayList, ByVal mFlag As MailEventFlag)

        ''
        '' 查詢IDB   - 因應全昆傳票起始或接收需要不顯示 , 需一開始得知是否為全昆    
        '' 將查詢IDB function移至此處
        ''           
        ''/*---------------  合併查詢品目字串 ----------------------
        'Dim query_string As String = String.Empty
        'Dim query_string_tp As String = String.Empty   ' TP 件使用

        'For Each Item As QuerySQL.RowCase In mData

        '    'If Not Item.Flag_Add Then
        '    '    Continue For
        '    'End If

        '    If Item.Mail_Title.Contains("TP測試") Then
        '        query_string_tp &= "*'" & Item.EIP_CODE_WIP & "'*"
        '    Else
        '        query_string &= "*'" & Item.EIP_CODE_WIP & "'*"
        '    End If

        'Next
        ''-----------------  合併查詢品目字串 ----------------------*/

        ''查詢非TP件
        'Dim arls_result As ArrayList = gSQL_Func.QueryIDB_Muti(mData, query_string)

        ''查詢 TP件
        'arls_result = gSQL_Func.QueryIDB_TP(arls_result, query_string_tp)



        '結果排序日期      
        '   arls_result = SortDateTime(arls_result)



        'Dim arls_TotResult As New ArrayList
        'If arls_result.Count > 0 Then

        '    For Each item As QuerySQL.RowCase In arls_result
        '        arls_TotResult.Add(item)
        '    Next
        '    ' arls_TotResult.AddRange(arls_result)

        'End If

        'If arls_result_TP.Count > 0 Then

        '    For Each item As QuerySQL.RowCase In arls_result_TP
        '        arls_TotResult.Add(item)
        '    Next

        '    'arls_TotResult.AddRange(arls_result_TP)
        'End If
        gObC_DATA.Clear()

        For Each item As QuerySQL.RowCase In arls_result

            If item.PROCESS_WIP Is Nothing Then
                Continue For
            End If

            If Not item.PROCESS_WIP.Equals("全昆") Then
                gObC_DATA.Add(item)
            End If

        Next


        Select Case mFlag
            Case MailEventFlag.INI_MAIL         '開啟程式初始化使用

                'For Each item As QuerySQL.RowCase In arls_result

                '    If item.PROCESS_WIP Is Nothing Then
                '        Continue For
                '    End If

                '    If Not item.PROCESS_WIP.Equals("全昆") Then
                '        gObC_DATA.Add(item)
                '    End If

                'Next

               ' gNoify.ShowBalloonTip(3000, "", "傳票接件程式已啟動", ToolTipIcon.Info)

            Case MailEventFlag.NEW_MAIL         '收到新mail 程式使用

                'For Each item As QuerySQL.RowCase In arls_result

                '    If item.PROCESS_WIP Is Nothing Then
                '        Continue For
                '    End If

                '    If Not item.PROCESS_WIP.Equals("全昆") Then

                '        gObC_DATA.Insert(0, item)
                '    End If

                'Next

                '有加入項目在顯示提示
                '   If arls_result.Count > 0 Then

                'ls_MailList.ScrollIntoView(ls_MailList.Items(0))         '更新資料顯示在最上方

                ' gNoify.ShowBalloonTip(3000, "", "接收到新傳票!", ToolTipIcon.Info)

                ' End If

        End Select

    End Sub
    '
    '   傳票列表排序時間
    '
    Private Function SortDateTime(arls_result As ArrayList) As ArrayList

        Dim arls_sort As New ArrayList

        For Each item As QuerySQL.RowCase In arls_result

            Dim mDateTime As DateTime = item.MailRecTime



            For i = 0 To arls_sort.Count - 1

                Dim tmp_data As QuerySQL.RowCase = arls_sort.Item(i)

                Dim cpDateTime As DateTime = tmp_data.MailRecTime

                If cpDateTime.CompareTo(mDateTime) > 0 Then


                End If


            Next


        Next



        Return arls_sort


    End Function


    '
    ' 檢查日期 刪除 信件
    ' (目前設定定 2 天)
    ' 輸入:  mail 轉換過資訊  , outlook Item  輸出: true / false 
    Private Function CheckDeleteMail(ByVal mData As QuerySQL.RowCase, ByVal mMailItem As Outlook.MailItem) As Boolean

        Try
            Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")

            '指向刪除收件夾
            Dim mDeleteFolder As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)

            Dim mFlag As Boolean = False

            '現在時間減2天
            Dim mOrgDate As DateTime = Now.AddDays(-2)

            '信件接收時間
            If mData.MailRecTime = Nothing Then

                mMailItem.Subject = "Delete Mail"
                mMailItem.Save()
                mMailItem.Move(mDeleteFolder)
                Return True

            End If

            Dim mCmpDate As DateTime = mData.MailRecTime

            ' /*------------------ 刪除過久信件 --------------------- 
            '  If mCmpDate < mOrgDate Then
            If DateTime.Compare(mCmpDate, mOrgDate) < 0 Then

                '標記  Delete Mail 移至刪除資料夾
                mMailItem.Subject = "Delete Mail"
                mMailItem.Save()
                mMailItem.Move(mDeleteFolder)

                'Dim mDeleteMail As Outlook.MailItem = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems).Items.Find("[Subject] = Delete Mail")
                'mDeleteMail.Delete()
                Return True
            End If
            '-------------------- 刪除過久信件 ---------------------*/
        Catch ex As Exception

        End Try


        Return False

    End Function
#Region "複製傳票信件到 [程式]傳票資料夾  (程式開啟)"

    '
    '  複製傳票信件到 [程式]傳票資料夾 
    ' 目前使用 程式開啟時呼叫
    Private Sub CopyMailToFolder()

        Try
            Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")
            Dim mInbox As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)  '指向收件夾

            Dim InboxItems As Outlook.Items = mInbox.Items 'mail信件 Items

            '/*------------------ 讀取主收件夾迴圈 -------------------------
            For Each Mailobject In InboxItems

                '非信件跳出 (可能為開會ITEM)
                If Not TypeOf Mailobject Is Outlook.MailItem Then
                    Continue For
                End If

                Mailobject = CType(Mailobject, Outlook.MailItem) '轉型MailItem

                '寄件者
                Dim mSender As String = Mailobject.SenderName

                '主旨
                Dim mTitle As String = Mailobject.Subject

                '接收時間
                Dim mRecTime As DateTime = Mailobject.ReceivedTime

                '  檢查時間   預設2天前信件
                Dim Today As DateTime = Now
                Dim Last2Day As DateTime = Today.AddDays(-2)
                '/*--------------  時間過久不檢查 -------------------
                If mRecTime < Last2Day Then
                    Continue For
                End If
                '---------------   時間過久不檢查 -------------------*/   



                '/*-------------- 篩選主旨與收件者為傳票  ---------------
                'If Not mSender.Contains("KT99_簽核主機") Then
                '    Continue For
                'End If
                Dim mFlag As Boolean = True
                If mTitle.Contains("聯絡傳票") Or mTitle.Contains("TP測試件") Then
                    mFlag = False
                End If

                If mTitle.Contains("追加數量") Then
                    mFlag = True
                End If

                If mFlag Then
                    Continue For
                End If
                '--------------- 篩選主旨與收件者為傳票  ---------------*/      


                '/*-------------  檢查是否重複信件 -------------------
                '複製傳票信件前檢查重複
                'If Not CheckRepreatMail(Mailobject, mInbox) Then

                Dim mCopiedItem As Outlook.MailItem = Mailobject.Copy
                Dim NewFolder As Outlook.Folder = mInbox.Folders.Item(gFolderName)
                mCopiedItem.Move(NewFolder)

                'End If
                '-------------  檢查是否重複信件 -------------------*/

            Next
            '-------------------- 讀取主收件夾迴圈 -------------------------*/

        Catch ex As System.Exception

        End Try

    End Sub

#End Region

#Region "複製時檢查程式用資料夾是否有重複信件 "


    '
    ' 複製傳票信件前檢查 是否重複
    '
    Private Function CheckRepreatMail(ByVal chk_mailObj As Outlook.MailItem, ByVal chk_Inbox As Outlook.Folder) As Boolean

        Dim ProgFolder As Outlook.Folder = chk_Inbox.Folders.Item(gFolderName)  '指向 [程式] 傳票收件夾
        Dim ProgItems As Outlook.Items = ProgFolder.Items 'mail信件 Items

        Dim mItemData As QuerySQL.RowCase = ParseMail(chk_mailObj)

        For Each mProgItem As Outlook.MailItem In ProgItems

            Dim mProgData As QuerySQL.RowCase = ParseMail(mProgItem)

            ' 將申請時間 + 傳票號碼+ mail收到時間 相同者判為重複信件 
            'If mItemData.EIP_CODE_WIP.Equals(mProgData.EIP_CODE_WIP) And
            '            mItemData.RegTime.Equals(mProgData.RegTime) And
            '            mItemData.MailRecTime.Equals(mProgData.MailRecTime) Then

            ' 將申請時間 + 傳票號碼 相同者判為重複信件        (目前有可能為更新資料傳票  僅接收時間不同)

            If mItemData.EIP_CODE_WIP.Equals(mProgData.EIP_CODE_WIP) And
                        mItemData.RegTime.Equals(mProgData.RegTime) Then
                Return True

            End If

        Next

        Return False

    End Function

#End Region

#Region "解析Mail 文件內容 Function "

    '
    ' 解析Mail 文件內容 將資料存入 StoreData
    '  輸入 : Outlook.MailItem 輸出: StoreData
    '
    Private Function ParseMail(mMailItem As Outlook.MailItem) As QuerySQL.RowCase


        Dim tmp_data As New QuerySQL.RowCase
        Try



            '讀取信件內文
            Dim recBody As String = mMailItem.Body.ToString()
            Dim arr_tmp() As String = recBody.Split(vbCrLf)



            'mail主旨  
            tmp_data.Mail_Title = mMailItem.Subject.ToString.Replace("CheckFlag--", String.Empty)

            'mail 接收時間
            tmp_data.MailRecTime = mMailItem.ReceivedTime.ToString

            For Each str_Line As String In arr_tmp

                Select Case True

                    Case str_Line.Contains("申請者")

                        Dim tmp_person As String = str_Line.Replace("申請者", "").Replace(":", "").Replace("：", "").Trim

                        '申請者:KS1000_莊育錡 
                        If tmp_person.Contains("_") Then
                            tmp_person = tmp_person.Split("_")(1).Trim
                        End If
                        tmp_data.CREATE_EMP_WIP = tmp_person

                    Case str_Line.Contains("表單單號")

                        tmp_data.EIP_CODE_WIP = str_Line.Replace("表單單號:", "").Replace(":", "").Replace("：", "").Trim

                    Case str_Line.Contains("申請時間:")

                        tmp_data.RegTime = str_Line.Replace("申請時間:", "").Replace(":", "").Replace("：", "").Trim

                    Case str_Line.Contains("客戶") And str_Line.Contains("料號")

                        ' ex: 客戶：台灣國際航電股份有限公司 料號：ZE12537-05 數量：3000
                        Dim mIndex As Integer = str_Line.IndexOf("料號") '尋找料號字串位置

                        If mIndex = 0 Then
                            Continue For
                        End If

                        '取得客戶
                        Dim mCustomer As String = str_Line.Substring(0, mIndex).Replace("：", "").Replace(":", "").Replace("客戶", "").Trim
                        tmp_data.CUSTOMER_WIP = mCustomer


                        Dim mPn As String = str_Line.Substring(mIndex, str_Line.Length - mIndex) '去除客戶名稱

                        mIndex = mPn.IndexOf("數量") '尋找數量字串位置

                        If mIndex = 0 Then
                            Continue For
                        End If

                        '數量
                        Dim mQty As String = mPn.Substring(mIndex, mPn.Length - mIndex)
                        mQty = mQty.Replace("：", "").Replace("數量", "").Trim
                        tmp_data.COUNT_WIP = mQty

                        '料號
                        mPn = mPn.Substring(0, mIndex)  '去除數量文字
                        mPn = mPn.Replace("：", "").Replace("料號", "").Trim
                        tmp_data.ITEM_WIP = mPn

                    Case str_Line.Contains("連結表單")

                        'Hyperlink "https://eip.flexium.com.tw/Portal/common/redirect/default.aspx?r=O2omSWYfOIyACOx2Xu7wGNLyiXKyb2Vp6na9okeHQWZEETQRcgYJ2kXFJjSQfACb4-Y4zzbmPmEd0PIZuThoR7PxVLqE8BbVc1kNf8ieB8HNAvrbULOpu6Y6-Vu7ZlO23N_KRcO0CorWyQ5EqgyhONtY5Jvyc40ND-cTe_kprOwHTupuVZUYOQ!!"請點選此處， 連結表單！

                        Dim hyper_link As String = str_Line.Replace(Chr(34), "") '去除雙引號
                        hyper_link = hyper_link.Replace("HYPERLINK", "") '去除Hyperlink 字串
                        hyper_link = hyper_link.Replace("請點選此處，連結表單！", "").Trim  '去除Hyperlink 字串
                        tmp_data.Hyper_Link = hyper_link

                End Select

            Next

        Catch ex As Exception

        End Try

        Return tmp_data

    End Function

#End Region

#Region "判斷程式用資料夾是否存在"
    '
    '  判斷程式用資料夾是否存在   (目前暫定:  [程式]傳票資料夾) 若無則新增
    '
    Private Sub IniMailFolder()



        Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")
        Dim mInbox As Outlook.Folder = Nothing
        mInbox = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)  '指向收件夾      

        '可指向其他自訂資料夾 - 待討論
        ' mInbox = mNSpace.Folders.Item("封存")


        Dim FolderExistFlag As Boolean = False

        For Each item As Outlook.Folder In mInbox.Folders
            If item.Name.Equals(gFolderName) Then
                FolderExistFlag = True
            End If
        Next


        If Not FolderExistFlag Then
            mInbox.Folders.Add(gFolderName, Outlook.OlDefaultFolders.olFolderDrafts)  '新增子資料夾
        End If




    End Sub
#End Region

    '
    '  接收新mail 事件
    '
    Private Sub NewMailEvent()

        Dim mNSpace As Outlook.NameSpace = gOlApp.GetNamespace("MAPI")
        Dim mInbox As Outlook.Folder = mNSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)  '指向收件夾

        Dim mItems As Outlook.Items = mInbox.Items 'mail信件 Items

        '調整僅收取未讀信件
        mItems = mItems.Restrict("[Unread] = true")

        For Each Mailobject In mItems

            If Not TypeOf Mailobject Is Outlook.MailItem Then
                Continue For
            End If

            '寄件者
            Dim mSender As String = Mailobject.SenderName

            '主旨
            Dim mTitle As String = Mailobject.Subject

            '/*-------------- 篩選主旨與收件者為傳票  ---------------
            '主旨與收件者為傳票再接收
            'If Not mSender.Contains("KT99_簽核主機") Then
            '    Continue For
            'End If
            Dim mFlag As Boolean = True
            If mTitle.Contains("聯絡傳票") Or mTitle.Contains("TP測試件") Then

                mFlag = False
            End If

            If mTitle.Contains("追加數量") Then
                mFlag = True
            End If

            If mFlag Then
                Continue For
            End If
            '--------------- 篩選主旨與收件者為傳票  ---------------*/  



            '初始化  "[程式]傳票資料夾
            IniMailFolder()


            '/*-------------  檢查是否重複信件 -------------------
            '複製傳票信件前檢查重複
            ' If Not CheckRepreatMail(Mailobject, mInbox) Then

            Dim mCopiedItem As Outlook.MailItem = Mailobject.Copy
            Dim NewFolder As Outlook.Folder = mInbox.Folders.Item(gFolderName)
            mCopiedItem.Move(NewFolder)

            'End If
            '-------------  檢查是否重複信件 -------------------*/


        Next


        'Update 列表
        UpdateProgInbox(MailEventFlag.NEW_MAIL)

        mInbox = Nothing
        mItems = Nothing


    End Sub

    Dim gArls As New ArrayList
    Private Sub StackPanel_Loaded(sender As Object, e As RoutedEventArgs)

        gOlApp = New Outlook.Application

        gObC_DATA = New ObservableCollection(Of QuerySQL.RowCase)


        Dim gen As Random = New Random()
        For i = 0 To 10
            'Dim dt As Date
            'dt = RandomDate("01/01/1975")
            'System.Diagnostics.Debug.Write(FormatDateTime(dt, DateFormat.ShortDate).ToString)

            'Dim TimeSpan As TimeSpan = endDate - startDate
            'Dim newSpan As TimeSpan = New TimeSpan(0, randomTest.Next(0, (Int())timeSpan.TotalMinutes), 0)
            'Dim newDate As DateTime = startDate + newSpan


            Dim range As Integer = 5 * 365
            Dim RandomDate As DateTime = DateTime.Today.AddDays(-gen.Next(range))

            Debug.WriteLine(RandomDate)

            Dim rc As New QuerySQL.RowCase
            rc.MailRecTime = RandomDate

            gArls.Add(rc)

        Next

        Dim myComparer As IComparer = New DateCompare
        gArls.Sort(myComparer)


        AddHandler bgw.DoWork, AddressOf Run_Update_Mailbox
        AddHandler bgw.ProgressChanged, AddressOf bgw_ProgressChanged
        AddHandler bgw.RunWorkerCompleted, AddressOf bgw_RunWorkerCompleted
        bgw.WorkerReportsProgress = True



        bgw.RunWorkerAsync()



        Try
            '     Init_Mail_Event()
        Catch ex As Exception

        End Try


        Console.WriteLine()



        'Array.Sort(gArls, AddressOf QuerySQL.RowCase 




    End Sub

    Private Sub bgw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        Throw New NotImplementedException()
    End Sub

    Private Sub bgw_ProgressChanged(sender As Object, e As ProgressChangedEventArgs)

        Dim mData As ArrayList = e.UserState

    End Sub

    Dim bgw As New BackgroundWorker
    '
    ' 
    '
    Private Sub Run_Update_Mailbox(sender As Object, e As DoWorkEventArgs)


        Do
            UpdateProgInbox(MailEventFlag.INI_MAIL)

            bgw.ReportProgress(1, numbers.ToList < Int() > ());
            Threading.Thread.Sleep(5000)


        Loop

    End Sub

    Class DateCompare
        Implements IComparer
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

            'Dim dt1 As DateTime = x.MailRecTime
            Dim row1 As QuerySQL.RowCase = x
            Dim row2 As QuerySQL.RowCase = y
            'Dim dt2 As DateTime = y.MailRecTime

            'Compare the time components.
            'Dim result As Integer = dt1.TimeOfDay.CompareTo(dt2.TimeOfDay)

            Return row1.MailRecTime.CompareTo(row2.MailRecTime)

            'If result = 0 Then
            '    'The times are the same so compare the date components.
            '    result = dt1.Date.CompareTo(dt2.Date)
            'End If

            'Return result
        End Function


    End Class


    Public Function RandomDate(ByVal StartDate) As Date
        'returns random date between start date and now

        If Not IsDate(StartDate) Then Exit Function
        Dim dt = CDate(StartDate)
        Dim iDifferential = DateDiff(DateInterval.Day,
               dt, System.DateTime.Now)

        iDifferential = New Random(System.DateTime.Now.Millisecond).Next(0, iDifferential)

        dt = DateAdd(DateInterval.Day, iDifferential, dt)

        Return dt

    End Function
End Class

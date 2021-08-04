Imports System.IO
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Module Module1
    Public WEBECOMDBCONN As String = "Data Source=192.168.100.83;Initial Catalog=ECOMLIVEREPL;UID=xxx;PWD=xxxxx"
    Public WEBDBCONN As String = "Data Source=192.168.100.83;Initial Catalog=WEBLIVE;UID=xxx;PWD=xxxxx"
    Public ECOMDBCONN As String = "Data Source=ecom-db1;Initial Catalog=ECOMLIVE;UID=xxx;PWD=xxxxx"
    Public LargeImages As List(Of String)
   Public SmallImages As List(Of String)
   Sub Main()
      Dim googlefile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\GoogleFeed." & Now.ToString("yyyyMMddHHmm") & ".txt"
      Dim shopzillafile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\ShopZillaFeed." & Now.ToString("yyyyMMddHHmm") & ".txt"
        Dim shoppingfile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\EcommerceECN.txt"
        Dim ItemList As New List(Of MarketableItem)
      ItemList = GetMarketableItems()
      Using feedfile As New StreamWriter(googlefile)
         feedfile.AutoFlush = True
         feedfile.WriteLine("id" & vbTab & "title" & vbTab & "link" & vbTab & "price" & vbTab & "description" & vbTab & "condition" & vbTab & "brand" & vbTab & "mpn" & vbTab & "image_link" & vbTab & "google_product_category" & vbTab & "quantity" & vbTab & "availability" & vbTab & "online_only" & vbTab & "manufacturer" & vbTab & "expiration_date" & vbTab & "featured_product" & vbTab & "year" & vbTab & "c:keywords" & vbTab & "sale_price_effective_date" & vbTab & "tax" & vbTab & "shipping" & vbTab & "product_type" & vbTab & "age_group" & vbTab & "gender" & vbTab & "color" & vbTab & "size" & vbTab & "adwords_redirect" & vbTab & "adwords_labels")
         For Each m As MarketableItem In ItemList
            If Not String.IsNullOrEmpty(m.Brand) And Not m.imagelink.EndsWith("/") Then
               feedfile.WriteLine(m.Edpno & vbTab & m.ShortDesc & vbTab & m.httplink & vbTab & m.Price.ToString("f2") & vbTab & m.LongDescription & vbTab & "New" & vbTab & m.Brand & vbTab & m.Edpno & vbTab & m.imagelink & vbTab & m.GoogleCategory & vbTab & m.Quantity & vbTab & m.Availablity & vbTab & "y" & vbTab & m.Brand & vbTab & Now.AddMonths(1).ToString("yyyy-MM-ddTHH:mm:ss") & vbTab & "" & vbTab & Now.Year & vbTab & m.WebCategories & vbTab & m.EffectiveDate & vbTab & "US:CA:8.25:n" & vbTab & m.Shipping & vbTab & m.WebCategories & vbTab & "Adult" & vbTab & m.Gender & vbTab & "Multi Color" & vbTab & "One Size" & vbTab & m.httplink & vbTab & m.WebCategories)
            End If
         Next
      End Using
      Using feedfile As New StreamWriter(shopzillafile)
         feedfile.AutoFlush = True
         feedfile.WriteLine("Category ID" & vbTab & "Manufacturer" & vbTab & "Title" & vbTab & "Description" & vbTab & "Product URL" & vbTab & "Image URL" & vbTab & "SKU" & vbTab & "Availability" & vbTab & "Condition" & vbTab & "Ship Weight" & vbTab & "Ship Cost" & vbTab & "Bid" & vbTab & "Promotional Code" & vbTab & "UPC" & vbTab & "Price")
         For Each m As MarketableItem In ItemList
            If Not String.IsNullOrEmpty(m.Brand) And Not m.imagelink.EndsWith("/") Then
               feedfile.WriteLine(m.ShopZillaCategory & vbTab & m.Brand & vbTab & m.ShortDesc & vbTab & m.ShopZillaLongDescription & vbTab & m.httplink & vbTab & m.imagelink & vbTab & m.ItemNo & vbTab & m.ShopZillaAvailablity & vbTab & "New" & vbTab & "" & vbTab & m.Shipping.Replace("US::Ground:", "") & vbTab & "" & vbTab & "39" & vbTab & "" & vbTab & m.Price)
            End If
         Next
      End Using
      Using feedfile As New StreamWriter(shoppingfile)
         feedfile.AutoFlush = True
         feedfile.WriteLine("Unique Merchant SKU" & vbTab & "Product Name" & vbTab & "Product URL" & vbTab & "Image URL" & vbTab & "Current Price" & vbTab & "Stock Availability" & vbTab & "Condition" & vbTab & "Shipping Rate" & vbTab & "Brand / Manufacturer" & vbTab & "Product Description" & vbTab & "Product Type" & vbTab & "Category" & vbTab & "Category ID" & vbTab & "Estimated Ship Date" & vbTab & "Gender" & vbTab & "Age Range")
         For Each m As MarketableItem In ItemList
            If Not String.IsNullOrEmpty(m.Brand) And Not m.imagelink.EndsWith("/") Then
               feedfile.WriteLine(m.ItemNo & vbTab & m.ShortDesc & vbTab & m.httplink & vbTab & m.imagelink & vbTab & m.Price & vbTab & m.ShopZillaAvailablity & vbTab & "New" & vbTab & m.Shipping.Replace("US::Ground:", "") & vbTab & m.Brand & vbTab & m.ShoppingLongDescription & vbTab & m.WebCategories & vbTab & "Sports and Outdoors > Sport and Outdoor" & vbTab & "96424" & vbTab & "Same day" & vbTab & m.Gender & vbTab & "Adult")
            End If
         Next
      End Using
      Dim FeedDirectory As String = "\\USA2\Public\feed\" & Now.ToString("yyyyMMddHHmm") & "\"
      Dim report As String = ""
      If Not Directory.Exists(FeedDirectory) Then
         Directory.CreateDirectory(FeedDirectory)
      End If
      If Directory.Exists(FeedDirectory) Then
         For Each s As String In {googlefile, shoppingfile, shopzillafile}
            Dim fi As New FileInfo(s)
            Try
               fi.CopyTo(FeedDirectory & fi.Name, True)
               Select Case fi.Name.Substring(0, 3)
                  Case "Goo"
                     report &= "Google Feed:       " & FeedDirectory & fi.Name & vbCrLf
                  Case "Pri"
                     report &= "Shopping.com Feed: " & FeedDirectory & fi.Name & vbCrLf
                  Case "Sho"
                     report &= "Shopzilla Feed:    " & FeedDirectory & fi.Name & vbCrLf
               End Select
            Catch ex As Exception
               report &= "Failed to move " & s & " to Feed Directory" & vbCrLf
            End Try
         Next
      Else
      End If
      Dim msg As New MailMessage
      With msg
            .From = New MailAddress(My.Computer.Name & "@ecommerce.com")
            .To.Add("randy@ecommerce.com")
            .CC.Add("federico@ecommerce.com")
            .Subject = "Google, Shopzilla, Shopping.com Feed Files"
         '.Attachments.Add(New Attachment(googlefile))
         '.Attachments.Add(New Attachment(shopzillafile))
         '.Attachments.Add(New Attachment(shoppingfile))
         .Body = "Feed files saved in diretory: " & "file://" & FeedDirectory & vbCrLf & vbCrLf & report
      End With
      Dim smtp As New SmtpClient("usa2")
      Try
         smtp.Send(msg)
      Catch ex As Exception
         Dim smpt2 As New SmtpClient("barracuda")
         smpt2.Send(msg)
      End Try
   End Sub
   Function GetFileListing(ByVal WhichList As String) As List(Of String)
      Dim tmpResult As New List(Of String)
      Dim ftpclient As FtpWebRequest
      If WhichList = "SMALL" Then
            ftpclient = FtpWebRequest.Create("ftp://www.ecommerce.com/960x600/")
        Else
            ftpclient = FtpWebRequest.Create("ftp://www.ecommerce.com/1100x750/")
        End If
      ftpclient.Credentials = New NetworkCredential("liveprodimgs", "lveWP66?")
      ftpclient.Method = WebRequestMethods.Ftp.ListDirectory
      Using sr As New StreamReader(ftpclient.GetResponse().GetResponseStream)
         Dim line As String = sr.ReadLine
         While Not line Is Nothing
            tmpResult.Add(line.ToUpper)
            line = sr.ReadLine
         End While
      End Using
      Return tmpResult
   End Function
   Function GetMarketableItems() As List(Of MarketableItem)
      LargeImages = GetFileListing("LARGE")
      SmallImages = GetFileListing("SMALL")
      Dim tempResult As New List(Of MarketableItem)
      Dim QueryString As String =
"DECLARE @TableITM TABLE(EDPNO int,ITEMNO char (20),STYLE char (12),CATEGORY CHAR (04),STATUS char (02),PRICE Int,INVENTORY bigint,SHORTDESC VARCHAR(850),BRAND VARCHAR(100)) " & _
"INSERT INTO @TableITM " & _
"SELECT EDPNO,ITEMNO,STYLE,CATEGORY,STATUS,IC.PRICE,AVAILABLEINV,IW.DESCRIPTION_SORTABLE, " & _
"BRAND = (SELECT TOP 1 NAME FROM [WEBLIVE].[dbo].[CATEGORY] WHERE CATEGORY_ID IN " & _
"(SELECT V.CATEGORY_ID FROM [WEBLIVE].[dbo].[VIEW_CATEGORY_ITEM] V JOIN [WEBLIVE].[dbo].[CATEGORY_HIERARCHY] H ON V.CATEGORY_ID = H.CATEGORY_ID WHERE V.ITEM_EDP = IC.EDPNO AND PARENT_ID = 353)) " & _
"FROM ITEMMAST IC JOIN [WEBLIVE].[dbo].[ITEM_MAST] IW ON IC.EDPNO = IW.ITEM_EDP " & _
"WHERE ITEMNO IN (SELECT DISTINCT SUBSTRING(ITEMNO,1,10) AS STYLEID FROM ITEMMAST WHERE EDPNO IN (SELECT EDPNO FROM ITEMMAST WHERE STATUS IN ('A1','C2','C3','R1','TG'))) " & _
"AND STYLE != '' ORDER BY ITEMNO " & _
"INSERT INTO @TableITM " & _
"SELECT EDPNO,ITEMNO,STYLE,CATEGORY,STATUS,IC.PRICE,AVAILABLEINV,IW.DESCRIPTION_SORTABLE, " & _
"BRAND = (SELECT TOP 1 NAME FROM [WEBLIVE].[dbo].[CATEGORY] WHERE CATEGORY_ID IN " & _
"(SELECT V.CATEGORY_ID FROM [WEBLIVE].[dbo].[VIEW_CATEGORY_ITEM] V JOIN [WEBLIVE].[dbo].[CATEGORY_HIERARCHY] H ON V.CATEGORY_ID = H.CATEGORY_ID WHERE V.ITEM_EDP = IC.EDPNO AND PARENT_ID = 353)) " & _
"FROM ITEMMAST IC JOIN [WEBLIVE].[dbo].[ITEM_MAST] IW ON IC.EDPNO = IW.ITEM_EDP " & _
"WHERE SUBSTRING(ITEMNO,1,10) IN (SELECT DISTINCT SUBSTRING(ITEMNO,1,10) AS STYLEID FROM ITEMMAST WHERE EDPNO IN (SELECT EDPNO FROM ITEMMAST WHERE STATUS IN ('A1','C2','C3','R1','TG'))) " & _
"AND STYLE = '' AND STATUS IN ('A1','C2','C3','R1','TG') ORDER BY ITEMNO " & _
"INSERT INTO @TableITM " & _
"SELECT EDPNO,ITEMNO,STYLE,CATEGORY,STATUS,IC.PRICE,AVAILABLEINV,IW.DESCRIPTION_SORTABLE, " & _
"BRAND = (SELECT TOP 1 NAME FROM [WEBLIVE].[dbo].[CATEGORY] WHERE CATEGORY_ID IN " & _
"(SELECT V.CATEGORY_ID FROM [WEBLIVE].[dbo].[VIEW_CATEGORY_ITEM] V JOIN [WEBLIVE].[dbo].[CATEGORY_HIERARCHY] H ON V.CATEGORY_ID = H.CATEGORY_ID WHERE V.ITEM_EDP = IC.EDPNO AND PARENT_ID = 353)) " & _
"FROM ITEMMAST IC JOIN [WEBLIVE].[dbo].[ITEM_MAST] IW ON IC.EDPNO = IW.ITEM_EDP " & _
"WHERE STATUS IN ('K1','K2','J1','J2') AND (STOPSHOPDATE = '' OR STOPSHOPDATE = '00000000') " & _
"SELECT M.EDPNO,M.ITEMNO,M.STYLE,M.CATEGORY,M.STATUS,M.PRICE,M.INVENTORY,SHORTDESC,BRAND,NAME,DESCRIPTION_LONG " & _
"FROM @TableITM M JOIN [WEBLIVE].[dbo].[CATEGORY] C ON C.CATEGORY_ID IN (SELECT CATEGORY_ID FROM [WEBLIVE].[dbo].[CATEGORY_ITEM] WHERE ITEM_EDP = M.EDPNO) " & _
"JOIN [WEBLIVE].[dbo].[VIEW_ITEMMAST1] VM ON VM.EDPNO = M.EDPNO " & _
"WHERE M.EDPNO IN (SELECT ITEM_EDP FROM [WEBLIVE].[dbo].[CATEGORY_ITEM]) ORDER BY M.EDPNO"
      Dim addededpnos As New List(Of UInt64)
      Using conn As New SqlConnection(WEBECOMDBCONN)
         Dim cmd As New SqlCommand(QueryString, conn)
         Try
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader
            If r.HasRows Then
               While r.Read
                  Dim tmpItem As New MarketableItem(r)
                  If addededpnos.Contains(tmpItem.Edpno) Then
                     For Each i As MarketableItem In tempResult
                        If i.Edpno = tmpItem.Edpno Then
                           i.OurWebCategories.Add(tmpItem.OurWebCategories.Item(0))
                        End If
                     Next
                  Else
                     addededpnos.Add(tmpItem.Edpno)
                     tempResult.Add(tmpItem)
                  End If
               End While
            End If
         Catch ex As Exception
            Console.WriteLine(ex.Message)
         End Try
      End Using
      Return tempResult
   End Function
End Module
Public Class MarketableItem
   Public ItemNo As String = ""
   Public Edpno As UInt64 = 0
   Public Status As String = ""
   Public Style As String = ""
   Public Brand As String = ""
   Public ShortDesc As String = ""
   Public LongDesc As String = ""
   Public httplink As String = ""
   Public imagelink As String = ""
   Public Price As Decimal = 0
   Public AvailableQty As Int64 = 0
   Public KeyWords As String = ""
   Public ProductType As String = ""
   Public Color As String = ""
   Public Size As String = ""
   Private ESMCategory As String = ""
   Public AdWordsLabels As String = ""
   Public OurWebCategories As New List(Of String)
   Sub New(ByVal r As SqlDataReader)
      ItemNo = Trim(r.Item("ITEMNO"))
      Edpno = r.Item("EDPNO")
      Status = Trim(r.Item("STATUS"))
      Style = Trim(r.Item("STYLE"))
      Price = r.Item("PRICE") / 100
      ESMCategory = Trim(r.Item("CATEGORY"))
      ShortDesc = Trim(r.Item("SHORTDESC"))
      OurWebCategories.Add(Trim(r.Item("NAME")))
      If Not IsDBNull(r.Item("DESCRIPTION_LONG")) Then
         LongDesc = Trim(r.Item("DESCRIPTION_LONG"))
      End If
      If Status.StartsWith("W") Or Status.StartsWith("K") Or Status.StartsWith("J") Or ItemNo.StartsWith("999") Then
         AvailableQty = 10
      Else
         AvailableQty = r.Item("INVENTORY")
      End If
      If Not IsDBNull(r.Item("BRAND")) Then
         Brand = Trim(r.Item("BRAND"))
         Dim tmpProd As String = ShortDesc.Replace(" ", "-").Replace("'", "-").Replace(".", "-").Replace("/", "-").Replace("$", "")
         Dim finalprod As String = ""
         If tmpProd.Contains("---") Then
            finalprod = tmpProd
         ElseIf tmpProd.Contains("--") Then
            finalprod = tmpProd.Replace("--", "-")
         Else
            finalprod = tmpProd
         End If
            httplink = "http://www.ecommerce.com/Brand/" & Brand.Replace(" ", "-").Replace("'", "-").Replace(".", "-").Replace("/", "-").Replace("--", "-") & "/" & finalprod & ".axd"
        End If
      Dim tmpItemlink As String = ItemNo.Replace(" ", "_").Replace("/", "-") & "_0.jpg".ToUpper
      Select Case ItemNo.Substring(0, 3)
         Case "018", "175", "180"
            Dim tmpitm As String = tmpItemlink
            If Not LargeImages.Contains(tmpItemlink) Then
               tmpitm = FindClosestItem(tmpItemlink, LargeImages)
            End If
            tmpItemlink = tmpitm
                imagelink = "http://www.ecommerce.com/_productimages/1100x750/" & tmpItemlink
            Case Else
            Dim tmpitm As String = tmpItemlink
            If Not LargeImages.Contains(tmpItemlink) Then
               tmpitm = FindClosestItem(tmpItemlink, SmallImages)
            End If
            tmpItemlink = tmpitm
                imagelink = "http://www.ecommerce.com/_productimages/960x600/" & tmpItemlink
        End Select
   End Sub
   Private Function FindClosestItem(ByVal oneItem As String, ByVal anItemList As List(Of String)) As String
      Dim styleid As String = oneItem.Substring(0, 10)
      For Each s As String In anItemList
         If s.StartsWith(styleid) And s.Contains("_0.") Then
            Return s
         End If
      Next
      For Each s As String In anItemList
         If s.StartsWith(styleid) Then
            Return s
         End If
      Next
      Return ""
   End Function
   Public ReadOnly Property GoogleCategory As String
      Get
         Select Case ItemNo.Substring(0, 3)
            Case "002"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "005"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Handlebars"
            Case "010"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Bags & Panniers"
            Case "015"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Grips and Handlebar Tape"
            Case "015"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "025"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Water Bottles"
            Case "030"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Bottom Brackets"
            Case "031"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "035"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Brake Parts > Bicycle Brake Levers"
            Case "040"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Brake Parts"
            Case "048"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Brake Parts"
            Case "050"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Small Parts"
            Case "065"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheel Parts"
            Case "070"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Chains"
            Case "075"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts"
            Case "080"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Chainrings"
            Case "085"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Tools, Cleaners & Lubricants"
            Case "090"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts"
            Case "095"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Computers"
            Case "100"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Cranks"
            Case "104"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts"
            Case "105"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts"
            Case "108"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Chainrings"
            Case "115"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Brake Parts"
            Case "117"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "120"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "125"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "130"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Water Bottles"
            Case "135"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories"
            Case "145"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories"
            Case "155"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Tools, Cleaners & Lubricants"
            Case "160"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "165"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Forks"
            Case "170"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Forks"
            Case "175"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "180"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "185"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Tools, Cleaners & Lubricants"
            Case "185"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "185"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear > Bicycle Gloves"
            Case "185"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "190"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Grips and Handlebar Tape"
            Case "195"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Handlebars"
            Case "200"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Grips and Handlebar Tape"
            Case "205"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "210"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Computers"
            Case "215"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear > Bicycle Helmets"
            Case "220"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheel Parts > Bicycle Hubs"
            Case "225"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Lights & Reflectors"
            Case "230"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Locks"
            Case "235"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Mirrors"
            Case "240"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Drivetrain Parts > Bicycle Pedals"
            Case "245"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Pumps"
            Case "250"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Stands & Storage"
            Case "255"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories"
            Case "260"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheel Parts > Bicycle Wheel Rims"
            Case "270"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Saddles"
            Case "275"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Seatposts"
            Case "285"
               Return "Apparel & Accessories > Shoes > Athletic Shoes > Bicycle Shoes"
            Case "295"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts"
            Case "305"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Stems"
            Case "315"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Forks"
            Case "318"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Forks"
            Case "320"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Tires"
            Case "325"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Tools, Cleaners & Lubricants"
            Case "330"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Trainers"
            Case "340"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "343"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Tires"
            Case "343"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Tubes"
            Case "343"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheel Parts"
            Case "345"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Tubes"
            Case "363"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheel Parts > Bicycle Spokes"
            Case "365"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Parts > Bicycle Wheels"
            Case "370"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Accessories > Bicycle Stands & Storage"
            Case "400"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "401"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "402"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case "405"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "500"
               Return "Apparel & Accessories > Clothing"
            Case "502"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "505"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear"
            Case "506"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear"
            Case "510"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "516"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "520"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear"
            Case "525"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "526"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "527"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "530"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "535"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "536"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "537"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "540"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Tights"
            Case "545"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Shorts"
            Case "550"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Tights"
            Case "565"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "570"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "575"
               Return "Sporting Goods > Outdoor Recreation > Cycling > Bicycle Gear"
            Case "580"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "581"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "585"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Jerseys"
            Case "586"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear > Bicycle Shorts"
            Case "595"
               Return "Apparel & Accessories > Clothing > Activewear > Bicycle Activewear"
            Case "999"
               Return "Sporting Goods > Outdoor Recreation > Cycling"
            Case Else
               Return "Sporting Goods > Outdoor Recreation > Cycling"
         End Select
         Return "Sporting Goods > Outdoor Recreation > Cycling"
      End Get
   End Property
   Public ReadOnly Property ShopZillaCategory As String
      Get
         Select Case ItemNo.Substring(0, 3)
            Case "002"
               Return "100001449"
            Case "005"
               Return "100001449"
            Case "010"
               Return "100001449"
            Case "015"
               Return "100001449"
            Case "018"
               Return "100001448"
            Case "025"
               Return "100001449"
            Case "030"
               Return "100001449"
            Case "031"
               Return "100001449"
            Case "035"
               Return "100001449"
            Case "040"
               Return "100001449"
            Case "048"
               Return "100001449"
            Case "050"
               Return "100001449"
            Case "065"
               Return "100001449"
            Case "070"
               Return "100001449"
            Case "075"
               Return "100001449"
            Case "080"
               Return "100001449"
            Case "085"
               Return "100001449"
            Case "090"
               Return "100001449"
            Case "095"
               Return "100001449"
            Case "100"
               Return "100001449"
            Case "104"
               Return "100001449"
            Case "105"
               Return "100001449"
            Case "108"
               Return "100001449"
            Case "115"
               Return "100001449"
            Case "117"
               Return "100001449"
            Case "120"
               Return "100001449"
            Case "125"
               Return "100001449"
            Case "130"
               Return "100001449"
            Case "135"
               Return "100001449"
            Case "145"
               Return "100001449"
            Case "155"
               Return "100001449"
            Case "160"
               Return "100001449"
            Case "165"
               Return "100001449"
            Case "170"
               Return "100001449"
            Case "175"
               Return "100001448"
            Case "180"
               Return "100001448"
            Case "185"
               Return "100001449"
            Case "185"
               Return "100001449"
            Case "185"
               Return "100001449"
            Case "185"
               Return "100001449"
            Case "190"
               Return "100001449"
            Case "195"
               Return "100001449"
            Case "200"
               Return "100001449"
            Case "205"
               Return "100001449"
            Case "210"
               Return "100001449"
            Case "215"
               Return "100001449"
            Case "220"
               Return "100001449"
            Case "225"
               Return "100001449"
            Case "230"
               Return "100001449"
            Case "235"
               Return "100001449"
            Case "240"
               Return "100001449"
            Case "245"
               Return "100001449"
            Case "250"
               Return "100001449"
            Case "255"
               Return "100001449"
            Case "260"
               Return "100001449"
            Case "270"
               Return "100001449"
            Case "275"
               Return "100001449"
            Case "285"
               Return "100001449"
            Case "295"
               Return "100001449"
            Case "305"
               Return "100001449"
            Case "315"
               Return "100001449"
            Case "318"
               Return "100001449"
            Case "320"
               Return "100001449"
            Case "325"
               Return "100001449"
            Case "330"
               Return "100001449"
            Case "340"
               Return "100001449"
            Case "343"
               Return "100001449"
            Case "343"
               Return "100001449"
            Case "343"
               Return "100001449"
            Case "345"
               Return "100001449"
            Case "363"
               Return "100001449"
            Case "365"
               Return "100001449"
            Case "370"
               Return "100001449"
            Case "400"
               Return "100001449"
            Case "401"
               Return "100001449"
            Case "402"
               Return "100001449"
            Case "405"
               Return "100001449"
            Case "500"
               Return "100001450"
            Case "502"
               Return "100001450"
            Case "505"
               Return "100001450"
            Case "506"
               Return "100001450"
            Case "510"
               Return "100001450"
            Case "516"
               Return "100001450"
            Case "520"
               Return "100001450"
            Case "525"
               Return "100001450"
            Case "526"
               Return "100001450"
            Case "527"
               Return "100001450"
            Case "530"
               Return "100001450"
            Case "535"
               Return "100001450"
            Case "536"
               Return "100001450"
            Case "537"
               Return "100001450"
            Case "540"
               Return "100001450"
            Case "545"
               Return "100001450"
            Case "550"
               Return "100001450"
            Case "565"
               Return "100001450"
            Case "570"
               Return "100001450"
            Case "575"
               Return "100001450"
            Case "580"
               Return "100001450"
            Case "581"
               Return "100001450"
            Case "585"
               Return "100001450"
            Case "586"
               Return "100001450"
            Case "595"
               Return "100001450"
            Case "999"
               Return "100001449"
            Case Else
               Return "100001449"
         End Select
         Return "100001449"
      End Get
   End Property
   Public ReadOnly Property Quantity As Int64
      Get
         If AvailableQty > 0 Then
            Return AvailableQty
         Else
            Return 0
         End If
      End Get
   End Property
   Public ReadOnly Property Availablity As String
      Get
         If AvailableQty > 0 Then
            Return "in stock"
         Else
            Return "available for order"
         End If
      End Get
   End Property
   Public ReadOnly Property ShopZillaAvailablity As String
      Get
         If AvailableQty > 0 Then
            Return "In Stock"
         Else
            Return "Back-Order"
         End If
      End Get
   End Property
   Public ReadOnly Property EffectiveDate As String
      Get
         Return Now.ToString("yyyy-MM-dd") & "T01:00:09Z/" & Now.AddMonths(1).ToString("yyyy-MM-dd") & "T01:00:09Z"
      End Get
   End Property
   Public ReadOnly Property Shipping As String
      Get
         If ItemNo.StartsWith("999") Then
            Return "US::Ground:0.00"
         End If
         If Price <= 19.99 Then
            Return "US::Ground:5.99"
         ElseIf Price <= 49.99 Then
            Return "US::Ground:7.99"
         ElseIf Price <= 75.99 Then
            Return "US::Ground:9.99"
         ElseIf Price <= 99.99 Then
            Return "US::Ground:10.99"
         ElseIf Price <= 149.99 Then
            Return "US::Ground:12.99"
         ElseIf Price <= 199.99 Then
            Return "US::Ground:13.99"
         ElseIf Price <= 299.99 Then
            Return "US::Ground:14.99"
         Else
            Return "US::Ground:15.99"
         End If
      End Get
   End Property
   Public ReadOnly Property Gender As String
      Get
         If ESMCategory.Contains("W") Then
            Return "Female"
         Else
            Return "Unisex"
         End If
      End Get
   End Property
   Public ReadOnly Property WebCategories As String
      Get
         Dim tmpReasult As String = ""
         For Each s As String In OurWebCategories
            tmpReasult &= s & ","
         Next
         Return tmpReasult.Substring(0, tmpReasult.Length - 1)
      End Get
   End Property
   Public ReadOnly Property LongDescription As String
      Get
         Dim s As String = LongDesc.Replace(vbTab, " ").Replace(vbCr, " ").Replace(vbCrLf, " ").Replace(vbLf, " ")
         Dim reg As New Regex("\<[^\>]*\>")
         Dim tmpResult As String = reg.Replace(s, String.Empty)
         If String.IsNullOrEmpty(tmpResult.Trim) Then
            Return ShortDesc
         ElseIf tmpResult.Length >= 10000 Then
            Return tmpResult.Substring(0, 9990) & "..."
         Else
            Return tmpResult
         End If
      End Get
   End Property
   Public ReadOnly Property ShopZillaLongDescription As String
      Get
         Dim s As String = LongDesc.Replace(vbTab, " ").Replace(vbCr, " ").Replace(vbCrLf, " ").Replace(vbLf, " ")
         Dim reg As New Regex("\<[^\>]*\>")
         Dim tmpResult As String = reg.Replace(s, String.Empty)
         If String.IsNullOrEmpty(tmpResult.Trim) Then
            Return ShortDesc
         ElseIf tmpResult.Length >= 1000 Then
            Return tmpResult.Substring(0, 995) & "..."
         Else
            Return tmpResult
         End If
      End Get
   End Property
   Public ReadOnly Property ShoppingLongDescription As String
      Get
         Dim s As String = LongDesc.Replace(vbTab, " ").Replace(vbCr, " ").Replace(vbCrLf, " ").Replace(vbLf, " ")
         Dim reg As New Regex("\<[^\>]*\>")
         Dim tmpResult As String = reg.Replace(s, String.Empty)
         If String.IsNullOrEmpty(tmpResult.Trim) Then
            Return ShortDesc
         ElseIf tmpResult.Length >= 4000 Then
            Return tmpResult.Substring(0, 3995) & "..."
         Else
            Return tmpResult
         End If
      End Get
   End Property
End Class

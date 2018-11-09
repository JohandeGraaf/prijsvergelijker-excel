Attribute VB_Name = "Module1"
Private sht_prices As Worksheet
Private sht_settings As Worksheet
Private start_row As String
Private remove_price As Boolean

Private webshops As Scripting.Dictionary

Public Sub UpdatePrices()
    Setup
    
    For Each Key In webshops.Keys()
        If Key = "You-mobile" Then
            UpdatePriceYoumobile webshops(Key)
        ElseIf Key = "Belsimpel" Then
            UpdatePriceBelsimpel webshops(Key)
        ElseIf Key = "Pricewatch" Then
            UpdatePricePricewatch webshops(Key)
        End If
    Next
    
    sht_prices.Cells(3, 4).Value = Now
    Application.CalculateFull
End Sub

Private Sub Setup()
    Set sht_prices = ThisWorkbook.Sheets("Prijzen")
    Set sht_settings = ThisWorkbook.Sheets("Instellingen")
    start_row = sht_settings.Cells(16, 1).Value
    
    Dim rp As String: rp = sht_settings.Cells(16, 2).Value
    If IsNotEmptyString(rp) And (LCase(rp) = "ja" Or LCase(rp) = "yes") Then
        remove_price = True
    Else
        remove_price = False
    End If
    
    Dim shops As Scripting.Dictionary
    Set shops = New Scripting.Dictionary
    For i = 2 To sht_settings.Rows(Rows.Count).End(xlUp).Row
        If IsNotEmptyString(sht_settings.Cells(i, 1).Value) And Not IsError(sht_settings.Cells(i, 3)) And Not IsError(sht_settings.Cells(i, 5)) Then
            Dim shop As Scripting.Dictionary
            Set shop = New Scripting.Dictionary
            
            shop.Add "Name", sht_settings.Cells(i, 1).Value
            shop.Add "PriceCol", sht_settings.Cells(i, 3).Value
            shop.Add "LinkCol", sht_settings.Cells(i, 5).Value
            shop.Add "MetaCol", sht_settings.Cells(i, 7).Value
            
            shops.Add sht_settings.Cells(i, 1).Value, shop
        End If
    Next
    
    Set webshops = shops
End Sub


Private Sub UpdatePriceYoumobile(ByVal settings As Scripting.Dictionary)
    Dim api_youmobile As String
    api_youmobile = "https://you-mobile.nl/index.php?option=com_mijoshop&format=raw&tmpl=component&route=product/product/updatePrice"
    
    Dim link_col As Variant: link_col = settings("LinkCol")
    Dim price_col As Variant: price_col = settings("PriceCol")
    Dim meta_col As Variant: meta_col = settings("MetaCol")
    
    For i = start_row To sht_prices.Rows(Rows.Count).End(xlUp).Row
        Dim price As String
        price = ""
        
        If IsNotEmptyString(sht_prices.Cells(i, 1).Value) And IsNotEmptyString(sht_prices.Cells(i, link_col).Value) Then
            Dim product_id As String
            product_id = ""
            
            'Check if we have the product_id already
            Dim meta_inf As String
            meta_inf = sht_prices.Cells(i, meta_col).Value
            If IsNotEmptyString(meta_inf) Then
                Dim strArr() As String
                strArr = Split(meta_inf, ";##")
                If UBound(strArr) = 1 Then
                    If strArr(1) = sht_prices.Cells(i, link_col).Value Then
                        product_id = strArr(0)
                    End If
                End If
            End If
            
            'Get product_id from html if we don't have it already
            If Not IsNotEmptyString(product_id) Then
                Dim response As MSHTML.HTMLDocument
                Set response = GetHttpDoc(sht_prices.Cells(i, link_col).Value)
                
                If Not response Is Nothing Then
                    Dim el As MSHTML.HTMLHtmlElement
                    Set el = response.querySelector("input[type=hidden][name='product_id']")
                    
                    If Not el Is Nothing Then
                        If IsNotEmptyString(el.Value) Then
                            product_id = el.Value
                            sht_prices.Cells(i, meta_col).Value = el.Value + ";##" + sht_prices.Cells(i, link_col).Value
                        End If
                    End If
                End If
            End If
            
            'Call you-mobile api with product_id to get price
            If IsNotEmptyString(product_id) Then
                Dim responsePost As String
                responsePost = PostHttp(api_youmobile, "product_id=" + product_id + "&quantity=1")
                
                If IsNotEmptyString(responsePost) Then
                    Dim respObj As Object
                    Set respObj = ParseJson(responsePost)
                    
                    If Not respObj Is Nothing Then
                        If TypeName(respObj) = "Dictionary" And respObj.Exists("price") And respObj.Exists("special") Then
                            If TypeName(respObj("special")) = "Boolean" And Not respObj("special") Then
                                price = Trim(respObj("price"))
                            Else
                                price = Trim(respObj("special"))
                            End If
                            
                            Dim strStr: strStr = "€ "
                            Dim idx: idx = InStr(1, price, strStr)
                            If idx > 0 Then
                                price = Mid(price, idx + Len(strStr))
                                price = Replace(price, ",", "")
                                price = Replace(price, ".", "")
                                
                                If Len(price) > 1 Then
                                    price = Mid(price, 1, Len(price) - 2) + "." + Right(price, 2)
                                Else
                                    price = ".0" + price
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If IsNotEmptyString(price) Or remove_price Then
            sht_prices.Cells(i, price_col).Value = price
        End If
    Next
End Sub

Private Sub UpdatePriceBelsimpel(ByVal settings As Scripting.Dictionary)
    Dim url_belsimpel As String: url_belsimpel = "https://www.belsimpel.nl/?rnd=" + CStr(Rnd())
    Dim api_belsimpel_add As String: api_belsimpel_add = "https://www.belsimpel.nl/winkelwagen/json/add?rnd=" + CStr(Rnd())
    Dim api_belsimpel As String: api_belsimpel = "https://www.belsimpel.nl/popup_winkelwagen?rnd=" + CStr(Rnd())
    
    Dim link_col As Variant: link_col = settings("LinkCol")
    Dim price_col As Variant: price_col = settings("PriceCol")
    Dim meta_col As Variant: meta_col = settings("MetaCol")
    
    'Get csrf token
    Dim csrf_token As String
    Dim response_csrf As MSHTML.HTMLDocument
    Set response_csrf = GetHttpDoc(url_belsimpel)
    If Not response_csrf Is Nothing Then
        Dim el_csrf As MSHTML.HTMLHtmlElement
        Set el_csrf = response_csrf.querySelector("input[type=hidden][name=csrf_token]")
        
        If Not el_csrf Is Nothing Then
            If IsNotEmptyString(el_csrf.Value) Then
                csrf_token = el_csrf.Value
            End If
        End If
    End If
    
    For i = start_row To sht_prices.Rows(Rows.Count).End(xlUp).Row
        Dim price As String
        price = ""
        
        If IsNotEmptyString(sht_prices.Cells(i, 1).Value) And IsNotEmptyString(sht_prices.Cells(i, link_col).Value) Then
            Dim product_id As String
            product_id = ""
            
            'Check if we have the product_id already
            Dim meta_inf As String
            meta_inf = sht_prices.Cells(i, meta_col).Value
            If IsNotEmptyString(meta_inf) Then
                Dim strArr() As String
                strArr = Split(meta_inf, ";##")
                If UBound(strArr) = 1 Then
                    If strArr(1) = sht_prices.Cells(i, link_col).Value Then
                        product_id = strArr(0)
                    End If
                End If
            End If
            
            'Get product_id from html if we don't have it already
            If Not IsNotEmptyString(product_id) Then
                Dim response As MSHTML.HTMLDocument
                Set response = GetHttpDoc(sht_prices.Cells(i, link_col).Value)
                
                If Not response Is Nothing Then
                    'Check if we have csrf token
                    Dim el_csrf2 As MSHTML.HTMLHtmlElement
                    Set el_csrf2 = response.querySelector("input[type=hidden][name=csrf_token]")
                    
                    If Not el_csrf2 Is Nothing Then
                        If IsNotEmptyString(el_csrf2.Value) Then
                            csrf_token = el_csrf2.Value
                        End If
                    End If
                    
                    'Get productid
                    Dim el As MSHTML.HTMLHtmlElement
                    Set el = response.querySelector("input[type=hidden][name='product_ids[]']")
                    
                    If Not el Is Nothing Then
                        If IsNotEmptyString(el.Value) Then
                            product_id = el.Value
                            sht_prices.Cells(i, meta_col).Value = el.Value + ";##" + sht_prices.Cells(i, link_col).Value
                        End If
                    End If
                End If
            End If
            
            'Call belsimpel api with product_id to get price
            If IsNotEmptyString(product_id) Then
                Dim formData As String: formData = "pos=304&csrf_token=" + csrf_token + "&product_ids[]=" + product_id
                
                Dim responsePostAdd As String
                responsePostAdd = PostHttp(api_belsimpel_add, formData)
            
                Dim responsePost As String
                responsePost = PostHttp(api_belsimpel, formData)
                
                If IsNotEmptyString(responsePost) Then
                    Dim respObj As Object
                    Set respObj = ParseJson(responsePost)
                    
                    If Not respObj Is Nothing Then
                        If TypeName(respObj) = "Dictionary" And respObj.Exists("data") Then
                            Set respObj = respObj("data")
                        End If
                        If TypeName(respObj) = "Dictionary" And respObj.Exists("added_product") Then
                            Set respObj = respObj("added_product")
                        End If
                        If TypeName(respObj) = "Dictionary" And respObj.Exists("prices") Then
                            Set respObj = respObj("prices")
                        End If
                    
                        If TypeName(respObj) = "Dictionary" And respObj.Exists("sale_price") Then
                            price = Trim(respObj("sale_price"))
                            price = Replace(price, ",", "")
                            price = Replace(price, ".", "")
                            If Len(price) > 1 Then
                                price = Mid(price, 1, Len(price) - 2) + "." + Right(price, 2)
                            Else
                                price = ".0" + price
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If IsNotEmptyString(price) Or remove_price Then
            sht_prices.Cells(i, price_col).Value = price
        End If
    Next
End Sub

Private Sub UpdatePricePricewatch(ByVal settings As Scripting.Dictionary)
    Dim url_tweakers As String: url_tweakers = "https://tweakers.net/?rnd=" + CStr(Rnd())
    Dim url_tweakers_cookieacc As String: url_tweakers_cookieacc = "https://tweakers.net/my.tnet/cookies/?rnd=" + CStr(Rnd())

    Dim link_col As Variant: link_col = settings("LinkCol")
    Dim price_col As Variant: price_col = settings("PriceCol")
    Dim meta_col As Variant: meta_col = settings("MetaCol")
    
    'Check if we need to accept cookies
    Dim response_ck As MSHTML.HTMLDocument
    Set response_ck = GetHttpDoc(url_tweakers)
    
    If Not response_ck Is Nothing Then
        Dim el_token As MSHTML.HTMLHtmlElement
        Set el_token = response_ck.querySelector("input[type=hidden][name=tweakers_token]")
        Dim el_returnto As MSHTML.HTMLHtmlElement
        Set el_returnto = response_ck.querySelector("input[type=hidden][name=returnTo]")
        
        If Not el_token Is Nothing And Not el_returnto Is Nothing Then
            If IsNotEmptyString(el_token.Value) And IsNotEmptyString(el_returnto.Value) Then
                'accept cookies
                Dim response_ck_acc As String
                response_ck_acc = PostHttp(url_tweakers_cookieacc, "decision=accept&returnTo=" + el_returnto.Value + "&fragment=&tweakers_token=" + el_token.Value)
            End If
        End If
    End If
    
    For i = start_row To sht_prices.Rows(Rows.Count).End(xlUp).Row
        Dim price As String
        price = ""
        
        If IsNotEmptyString(sht_prices.Cells(i, 1).Value) And IsNotEmptyString(sht_prices.Cells(i, link_col).Value) Then
            Dim canonical_url As String
            canonical_url = ""
            
            'Check if we have the cannonical_url already
            Dim meta_inf As String
            meta_inf = sht_prices.Cells(i, meta_col).Value
            If IsNotEmptyString(meta_inf) Then
                Dim strArr() As String
                strArr = Split(meta_inf, ";##")
                If UBound(strArr) = 1 Then
                    If strArr(1) = sht_prices.Cells(i, link_col).Value Then
                        canonical_url = strArr(0)
                    End If
                End If
            End If
            
            'Get canonical_url from html if we don't have it already
            If Not IsNotEmptyString(canonical_url) Then
                Dim response_cn As MSHTML.HTMLDocument
                Set response_cn = GetHttpDoc(sht_prices.Cells(i, link_col).Value)
                
                If Not response_cn Is Nothing Then
                    'Get canonical_url
                    Dim el_cn As MSHTML.HTMLHtmlElement
                    Set el_cn = response_cn.querySelector("link[rel='canonical']")
                    
                    If Not el_cn Is Nothing Then
                        If IsNotEmptyString(el_cn.getAttribute("href")) Then
                            canonical_url = el_cn.getAttribute("href")
                            sht_prices.Cells(i, meta_col).Value = el_cn.getAttribute("href") + ";##" + sht_prices.Cells(i, link_col).Value
                        End If
                    End If
                End If
            End If
            
            'Get pricewatch site with canonical_url to get lowest price
            If IsNotEmptyString(canonical_url) Then
                Dim response As MSHTML.HTMLDocument
                Set response = GetHttpDoc(canonical_url + "?orderField=totalprice&orderSort=asc&rnd=" + CStr(Rnd()))
                
                If Not response Is Nothing Then
                    'Get lowest price from table
                    Dim el As MSHTML.HTMLHtmlElement
                    Set el = response.querySelector("#listing table tbody tr:first-child .shop-price p:first-child a:first-child")
                    
                    If Not el Is Nothing Then
                        If IsNotEmptyString(el.innerHTML) Then
                            price = Trim(Replace(el.innerHTML, "€ ", ""))
                            price = Replace(price, ",-", ",00")
                            price = Replace(price, ",", "")
                            price = Replace(price, ".", "")
                            
                            If Len(price) > 1 Then
                                price = Mid(price, 1, Len(price) - 2) + "." + Right(price, 2)
                            Else
                                price = ".0" + price
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If IsNotEmptyString(price) Or remove_price Then
            sht_prices.Cells(i, price_col).Value = price
        End If
    Next
End Sub

Private Function GetHttpDoc(ByVal URL As String) As MSHTML.HTMLDocument
    Dim response As String
    response = GetHttp(URL)
    
    Dim doc As Object
    If IsNotEmptyString(response) Then
        Set doc = New MSHTML.HTMLDocument
        doc.Open
        doc.Write response
        doc.Close
    Else
        Set doc = Nothing
    End If
    Set GetHttpDoc = doc
End Function

Private Function GetHttp(ByVal URL As String) As String
    On Error GoTo GetHttpErr
    
    Dim xhr As MSXML2.XMLHTTP60
    Set xhr = New MSXML2.XMLHTTP60
    With xhr
        .Open "GET", URL, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0"
        .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        .setRequestHeader "Accept-Language", "nl,en-US;q=0.7,en;q=0.3"
        .setRequestHeader "Accept-Encoding", "gzip, deflate"
        .setRequestHeader "Dnt", "1"
        .setRequestHeader "Connection", "keep-alive"
        .setRequestHeader "Upgrade-Insecure-Requests", "1"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Cache-Control", "no-cache"
        .Send
        
        If .readyState = 4 And (.Status = 200 Or .Status = 202) And IsNotEmptyString(.responseText) Then
            GetHttp = .responseText
        Else
            GetHttp = ""
        End If
    End With
    Exit Function
GetHttpErr:
    GetHttp = ""
End Function

Private Function PostHttp(ByVal URL As String, ByVal formData As String) As String
    On Error GoTo PostHttpErr
    
    Dim xhr As MSXML2.XMLHTTP60
    Set xhr = New MSXML2.XMLHTTP60
    With xhr
        .Open "POST", URL, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0"
        .setRequestHeader "Accept-Language", "nl,en-US;q=0.7,en;q=0.3"
        .setRequestHeader "Dnt", "1"
        .setRequestHeader "Connection", "keep-alive"
        .setRequestHeader "Upgrade-Insecure-Requests", "1"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Cache-Control", "no-cache"
        .Send formData
        
        If .readyState = 4 And (.Status = 200 Or .Status = 202) And IsNotEmptyString(.responseText) Then
            PostHttp = .responseText
        Else
            PostHttp = ""
        End If
    End With
    Exit Function
PostHttpErr:
    PostHttp = ""
End Function

Private Function ParseJson(ByVal json As String) As Object
    On Error GoTo ParseJsonErr
    
    Dim jsonObj As Object
    Set jsonObj = JsonConverter.ParseJson(json)

    Set ParseJson = jsonObj
    Exit Function
ParseJsonErr:
    Set ParseJson = Nothing
End Function

Private Function IsNotEmptyString(ByVal str As String) As Boolean
    If Len(Trim(str)) > 0 Then
        IsNotEmptyString = True
    End If
End Function

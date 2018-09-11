Attribute VB_Name = "gitHub_webScrap"
Sub web_Scraping_fill_in_login_page()

    Dim myId, myPass As String
    myId = "EMazarakis"
    myPass = "me1990#"


    MsgBox ("Connecting to GitHub site....")
    
    ' URL of the site where we want to visit
    Website = "https://github.com/login"
    
    'Set up the Internet Explorer object
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    
    With IE
        .Visible = True
        .navigate Website
    End With
    
    
    ' Wait untill all the data has been loaded for the IE object
    Do
        DoEvents
    Loop Until IE.readystate = READYSTATE_COMPLETE
    
    
    ' This is the web page, i.e the html code
    Dim myDocument As HTMLDocument
    Set myDocument = IE.document

  
    ' Wait a few seconds, before you paste the username and password
    Application.Wait (Now + TimeValue("0:00:05"))
  
    Dim id_User, pass_User As Object                        ' Represents id_User, pass_User & btnLogin as object
    Set id_User = myDocument.getElementById("login_field")  ' Returns an HTML element which has Id == login-email
    Set pass_User = myDocument.getElementById("password")   ' Returns an HTML element which has Id == login-password
      
    id_User.Value = myId        ' Set the value of the EMAIL at the right field
    pass_User.Value = myPass    ' Set the value of the PASSWORD at the right field


    ' Wait a few seconds, before you press the login button
    Application.Wait (Now + TimeValue("0:00:03"))

    
    ' Represents the Login Button
    Dim btnLogin As Object
    Dim tag_Collections As Object
    Set tag_Collections = myDocument.getElementsByTagName("input") ' list of all <input> tags of the html document
    

    ' For each tag <input>, check and find the login Button
    For Each elem In tag_Collections

        ' I am searching in each <input> tag which one has the "value" attribute equal to "Sign in"
        If elem.getAttribute("value") = "Sign in" Then
        
            MsgBox ("Found the login button")
            Set btnLogin = elem             ' I am keeping the right one <a> tag. because this is the logout button
            Exit For
            
        End If
        
    Next elem       'Go to next tag of the collection
    
    ' Click the login button
    btnLogin.Click
    
    
    ' Wait a few seconds in the home page before you exit
    Application.Wait (Now + TimeValue("0:00:15"))
    
    
    '-------------------------------------------------
    
    ' //TODO : Do Whatever you want in the GitHub site
    
    '-------------------------------------------------
    
    
    ' Represents the Logout Button
    Dim logout_Button As Object
    Dim a_tag_Collections As Object
    Set a_tag_Collections = myDocument.getElementsByClassName("dropdown-item dropdown-signout")
    

    ' For each tag with class name equal to "dropdown-item dropdown-signout", check and find the logout Button
    For Each elemm In a_tag_Collections

        'MsgBox (elemm.innerText)
        If elemm.innerText = "Sign out " Then
        
            MsgBox ("Found the logout button")
            Set logout_Button = elemm            ' I am keeping the right one tag, because this is the logout button
            Exit For
        End If
        
    Next elemm       'Go to next tag of the collection
    

    ' Click the logout button
    logout_Button.Click
    
    MsgBox ("You are out the Web page.")
    

End Sub




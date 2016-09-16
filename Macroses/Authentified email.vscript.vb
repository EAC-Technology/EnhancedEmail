'--- General libs
'#include(lib_lang)
'#include(lib_utils)
'#include(lib_quicksort)
'#include(lib_func_parser_dialog)
'#include(lib_dialog_2)
'#include(lib_EAC_functions)


'Initial SetUp
args = xml_dialog.get_answer
logger ("Data sent :" & tojson(args))
logger Appinmail.utils.parseEmails(args)

set current_user = ProAdmin.current_user

if "step" in args then
  if args( "step" )="" then
    operation_step ="Init"
  else
    operation_step = args( "step" )
  end if
else
  operation_step ="Init"
end if

'###############################################################################################
'#                                Screen Definition                                            #
'###############################################################################################
'VDOMData = DynamicVDOM.render(vdomxml) 
set MyForms = New XMLDialogBuilder

'-------------------------------------- VDOM Screen TEST ---------------------------------------
' TEST to include VDOM XML in XML Dialog
'-----------------------------------------------------------------------------------------------    
    MyForms.addScreen("Show VDOM")
    MyForms.Screen("Show VDOM").addComponent("VDOMXML","hypertext")
    
    vdomxml =  "<TEXT name='test' value='test'/>"
   
    DynamicVDOM.left="0"
    DynamicVDOM.top="0"
    DynamicVDOM.width="600"
    DynamicVDOM.width="300"
    'VDOMData = DynamicVDOM.render(vdomxml) 
    'MyForms.Screen("Show VDOM").Component("VDOMXML").value(VDOMData)

    'logger Replace(VDOMData, "<", "&lt;")
    
'-------------------------------------- Std Composer -------------------------------------------
' Main error Window
'-----------------------------------------------------------------------------------------------    
  MyForms.addScreen("Std Composer")
  MyForms.Screen("Std Composer").width  = 800
  MyForms.Screen("Std Composer").Height = 800
  MyForms.Screen("Std Composer").Title  = "Standard mail composer"
  'MyForms.Screen("Std Composer").addComponent("fromemail","livesearch")
  'MyForms.Screen("Std Composer").Component("fromemail").label("From :")
  MyForms.Screen("Std Composer").addComponent("toemail","livesearch")
  MyForms.Screen("Std Composer").Component("toemail").label("To :")
  MyForms.Screen("Std Composer").addComponent("subject","TextBox")
  MyForms.Screen("Std Composer").Component("subject").label("Subject :")
  MyForms.Screen("Std Composer").addComponent("message","RichTextArea")
  MyForms.Screen("Std Composer").Component("message").label("Message :")
  MyForms.Screen("Std Composer").Component("message").height("400")
  MyForms.Screen("Std Composer").addComponent("attach","FileUpload")
  MyForms.Screen("Std Composer").Component("attach").label("Attachments :")
  
  MyForms.Screen("Std Composer").addComponent("btns","btngroup")
  MyForms.Screen("Std Composer").Component("btns").addBtn("Send","sendEmail")
  MyForms.Screen("Std Composer").Component("btns").addBtn("Cancel","Exit")

'-------------------------------------- Msg Email sent -----------------------------------------
' Msg box to show the email was sent
'-----------------------------------------------------------------------------------------------
    MyForms.addScreen("Email Sent Msg")
    MyForms.Screen("Email Sent Msg").addComponent("comment","Text")
    MyForms.Screen("Email Sent Msg").Component("comment").setCenter(true)
    MyForms.Screen("Email Sent Msg").Component("comment").value("<br>Your email was correctly sent.")
    MyForms.Screen("Email Sent Msg").Title = "Sending email ..."
    MyForms.Screen("Email Sent Msg").addComponent("timeout","Timer")
    MyForms.Screen("Email Sent Msg").Component("timeout").setTimer("Exit",1000)
    
     
    
 Function tt(e, mailbox)
 
 content = args("message")
 
   ' userguid = Proadmin.currentuser().guid
   ' logger(userguid)
    
    vdomxml = "<CONTAINER name=""container1"" designcolor=""84A2F0"" top=""9"" height=""443"" width=""561"" left=""18"">"&_
    "<RICHTEXT name=""richtext2"" top=""75"" left=""37""/>"&_
      "<HYPERTEXT name=""text1"" left=""15"" top=""50"" width=""500"" height=""400"">"&_
        "<Attribute Name=""htmlcode"">"&_
            "<![CDATA[" & content & "]]>" &_
        "</Attribute>"&_
      "</HYPERTEXT>"&_
    "</CONTAINER>"
      
    logger(vdomxml)
    
    'e.dynamic = true               ' true/false whether it's dynamic content (boolean)
    e.auth = "internal"             ' internal/external authentication (string)
    e.session_token = ""         ' session token (string)
    e.login_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"   ' ID of the container where the login                             ' method is (string)
    e.login_method = "login"     ' name of the login method (string)
   ' e.get_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"     ' ID of the container where to call the method
                                ' to get data (string)
    'e.get_method = "call_macro"         ' name of the method to get data (string)
   ' e.get_data = "{""plugin_guid"": ""5d525ea5-bb1d-4eab-b5c3-77a30b5ebe67"", ""async"": 0, ""data"": {""guidsender"": """ + userguid +""", ""type"" : ""authentified""}, ""name"": ""loaddata""}"          ' data to be sent as pattern for GET request (string)
    'e.post_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"    ' ID of the container where to call the method
                                ' to post data (string)
    'e.post_method = "call_macro"       ' name of the method to post data (string)
   ' e.post_data = "{""plugin_guid"": ""5d525ea5-bb1d-4eab-b5c3-77a30b5ebe67"", ""async"": 0, ""data"": {""guidsender"": """+ userguid +""", ""type"" : ""authentified""}, ""name"": ""loaddata""}"        ' data to be sent as pattern for POST
                  
    e.api_server = Appinmail.utils.currentHost()      ' URL of backend server to handle EAC business
                                ' logic (string)
    e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"            ' ID of the VDOM application running the business
                                ' logic (string)
    e.events_data = ""        ' JSON structure that represents events (string)
    e.vdomxml_data = vdomxml      ' XML that represents the VDOM XML for this
                                ' object (string)
    'e.item_plugin = "itemplugin"    ' ID of the plugin to process the events of
                                ' interactive mail list item (string)
    'e.item_vdomxml = "itemvdomxml"  ' custom VDOM XML to render mail list item
                                ' for this email (string)
    'e.add_tag("tag1", "ff0000")     ' method to add tag with name and color (strings)
    'e.add_tag("tag2", "00ff00")
    'e.remove_tag("tag2")            ' method to remove tag by name
    e.eac_token = generateguid()     ' EAC token (string)
    e.eac_method = "new"            ' EAC method (string, new/update/delete)
    logger "WHOLEXML:--"
    logger e.get_wholexml()                ' return composed XML as string
    'e.set_events(events)            ' method to set events data, accepts JSON string
                                ' or dictionary

    files = Dictionary
    logger("Attachments: " & args("attach"))
    if args("attach") <> "" then
        attach = split(args("attach"), ",")
        for each a in attach
'            logger("Attachment: " & a)
            Set f = xml_dialog.uploadedfile(a)
'            logger("- name: " & f.name)
            files(f.name) = f.data
            f.remove()
        next
'        for each f in files
'            logger(f)
'        next
    end if
    xml_dialog.clearfiles()

    toEmails = Appinmail.utils.parseEmails(args)
    for each email in toEmails
        eacviewer_url = e.get_eacviewer_url(Appinmail.utils.currentHost(), email)
    
        try
            query = Dictionary
            query("user_email") = email
            user_info = appinmail.getUser(query)
            login = user_info("user_login")
        catch
            login = email
        end try
                
        try
            e.send(mailbox, email, "", "", args("subject"), "Use promail or link " + eacviewer_url, files)
            if login <> email then
                e.send(mailbox, login, "", "", args("subject"), "Use promail or link " + eacviewer_url, files)
            end if
        catch
        end try
    next
          
end function 
'############################################################################################
'#                                                                                          #
'#                         State machine start with state / Init /                          #
'#                                                                                          #
'############################################################################################

if instr(operation_step,">")<>0 then
  mainStep = split(operation_step,">")(0)
  subStep = split(operation_step,">")(1)
else
  mainStep = operation_step
  subStep = ""
end if

logger ("operation step :" & operation_step)

Select case mainStep

  case "Init" 
    MyForms.ShowScreen("Std Composer")

  case "sendEmail"
    
Set e = new EAC 
  Call tt(e, ProMail.selected_mailbox)
  
    MyForms.ShowScreen("Email Sent Msg")  

  case "Exit"
    logger ("End of Processing Rules")

end select
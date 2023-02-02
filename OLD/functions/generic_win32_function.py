# Win32com package (known as 'pywin32') provides access to Windows APIs
# Manipulate MS Office Apps, require installation 'pip install pywin32'
import win32com.client

# Start a MS Office App, can be run in the background or in front.
# Params : background - Run app in the background (BOOL)
#          application_codename - The app name to instantiate (STRING)
# Return : The started application object
def start_app(application_codename, background = True):
    if background == True:
        app_ = win32com.client.Dispatch(application_codename)
        app_.Visible = 0
    else:
        app_ = win32com.client.Dispatch(application_codename)
        app_.Visible = 1
    return app_

# Close a MS Office App
# Params : app_ - The application object to quit (OBJECT)
# Return : None
def exit_app(app_):
    app_.Quit()

# Close a document open in an MS Office App
# Params : doc_ - The document object to close (OBJECT)
# Return : None
def close_doc(doc_):
    doc_.Close()

# Save and close a document open in an MS Office App
# Params : doc_ - The document object to save & close (OBJECT)
#          location - Optionnaly SaveAs that location (STRING)
# Return : None
def save_doc(doc_, location=""):
    if location == "":
        doc_.Save()
    else:
        doc_.SaveAs(Filename:=location)

# Save doc and quit the application (all-in-one call)
# Params : app_ - The application object to quit (OBJECT)
#          doc_ - The document object to save & close (OBJECT)
#          location - Optionnaly SaveAs that location (STRING)
# Return : None
def save_doc_and_quit_app(app_, doc_, location=""):
    save_doc(doc_, location)
    close_doc(doc_)
    exit_app(app_)

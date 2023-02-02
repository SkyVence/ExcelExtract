import win32com.client

#https://stackoverflow.com/questions/67527415/how-to-connect-shapes-in-visio-using-vba-pythonpywin32

#Open visio using API from win32com
visio_app = win32com.client.Dispatch("Visio.Application")
#visio_app.Visible = 0
# OR
#visio_app = win32com.client.Dispatch("Visio.InvisibleApp")

#Select the visio template as Basic
visio_doc = visio_app.Documents.Add("Basic Diagram.vst")

#Selection 1st visio page
visio_page = visio_doc.Pages.Item(1)

#Select Stencils

visio_doc2 = visio_app.Documents.open("3015.vss")
#visio_doc2.close
visio_stencils = visio_app.Documents('3015.vss')

#Select Shape
shape_square = visio_stencils.Masters("ADM")

#Drop line (connector) & shapes at location X Y
line = visio_page.DrawLine(5, 4, 7, 4)  
shape1 = visio_page.Drop(shape_square, 4, 4)
shape2 = visio_page.Drop(shape_square, 8, 4)

#Get both ends of the line (connector) and attach them to the middle of the shapes
line_start = line.Cells("BeginX") 
line_end = line.Cells("EndX") 
line_start.GlueToPos(shape1, 0.5, 0.5)
line_end.GlueToPos(shape2, 0.5, 0.5)

#Do the same as above for a 3rd shape
line2 = visio_page.DrawLine(5, 4, 6, 5)  
line3 = visio_page.DrawLine(6, 5, 7, 4)  

shape_square2 = visio_stencils.Masters("15200")
shape3 = visio_page.Drop(shape_square2, 6, 5)

line2_start = line2.Cells("BeginX")
line2_end = line2.Cells("EndX") 
line2_start.GlueToPos(shape3, 0.5, 0.5)
line2_end.GlueToPos(shape2, 0.5, 0.5)

line3_start = line3.Cells("BeginX") 
line3_end = line3.Cells("EndX") 
line3_start.GlueToPos(shape3, 0.5, 0.5)
line3_end.GlueToPos(shape1, 0.5, 0.5)

# Save & quit
visio_doc.SaveAs('C:\\Users\\Adrien\\Desktop\\poc_visio.vsdx')

#visio_doc.close
#visio_doc2.close
visio_app.quit

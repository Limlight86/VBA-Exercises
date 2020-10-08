# Creating a VBA script to export a cells from Excel to Power Point and Outlook.

In order to get this script to work, we will need to have created a table already on excel. With this script we can create a 1 for 1 copy of any chosen cells in Excel and transfer them over to a new email in Outlook and a Power Point slide. Any previously created table or data set can be used.

To get started, lets open up Visual Basic in our chosen Excel sheet, the option is under the Developer header tab or can be opened by using the alt+f11 shortcut.

We will need to crate a new module. Using the Insert header tab, lets add a new module.

Inside the module we just created, we need to make sure the following libraries are imported:

```
Microsoft Excel Object Library 16.0
Microsoft Office Object Library 16.0
Microsoft Outlook Object Library 16.0
Microsoft PowerPoint Object Library 16.0
Microsoft Word Object Library 16.0
```

These libraries will contain much of the functionality that will be used later in the script when working with our Office apps.

## Create a new sub routine so we can populate with our code:

```
Sub CopyToPPandOutlook()

End Sub
```

Now we need to declare our variables from our object libraries that we will be using in the script, using the `Dim` variable statement:

```
    Dim PP As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.Slide

    Dim SlideTitle As String

    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector

    Dim oWordDoc As Word.Document
    Dim oWordRng As Word.Range
    Dim oWordTable As Word.Table
```

Remember that what we call these variables is up to you, use whatever names will flow well with your coding process.

For PowerPoint, we have included the ability to know what features and functionality the app has using the `PP` variable. We create a new presentation and a slide inside of it with `PPPres` and `PPSlide`. And finally, we have created an empty slide tittle to be populated further down our script with `SlideTitle`.

Similarly, for Outlook, we have created a variable to use the features of the Outlook app with `oLookApp`. We created a new email with `oLookItm` and finally we created the Inspector which is the window in Outlook that shows detailed information for a particular Outlook item. This is the window that displays when you double-click an item in an Outlook folder, this will be needed to edit our new email.

Lastly because Outlook is using Word under the hood for its text editing, we brought in the dependencies for Word, including the ability to select a range within the document and the ability to create a table.

## Creating our Outlook message

To create a copy of our table from Excel to Outlook we need to do the following steps:

Initializing our App and Email:
```
  Set oLookApp = New Outlook.Application

  Set oLookItm = oLookApp.CreateItem(olMailItem)
```

The following code will open the Outlook App on your pc (this step is not needed but its good to have all the tools at your disposal when working with these scripts):

` Shell ("OUTLOOK")`

Next, we need to do a series of commands and logic within a `With` statement, which executes a series of statements on a single object or a user-defined type.

```
  With oLookItm
    .To = "some@guy.com"
    .CC = "some@gal.com"
    .Subject = "Test Email"
    .Body = "Test table"

    .Display

    Set oLookIns = .GetInspector

    Set oWordDoc = oLookIns.WordEditor

    Set oWordRng = oWordDoc.Application.ActiveDocument.Content
        oWordRng.Collapse Direction:=wdCollapseEnd

    oWordRng.InsertBreak

    oWordRng.PasteAndFormat Worksheets("Excel2PP&OL").Range("C2:K23").CopyPicture
  End With
```

Let’s break down the following lines of code inside our newly created `With` statement that I have called `oLookItm` which in essence, will create the entire content of our new email.

```
  .To = "some@guy.com"
  .CC = "some@gal.com"
  .Subject = "Test Email"
  .Body = "Test table"       

  .Display
````

Here we have specified some basic parts of an email composition, the To, CC, Subject and a simple content for our email’s body. The last like `.Display` will actually show you in your UI the email being created.

```
  Set oLookIns = .GetInspector

  Set oWordDoc = oLookIns.WordEditor
```

These lines will have the inspector locate the Word part of the email, which we will be manipulating using our script to insert the table.

```
  Set oWordRng = oWordDoc.Application.ActiveDocument.Content
      oWordRng.Collapse Direction:=wdCollapseEnd

  oWordRng.InsertBreak
```

Using the above code, we are able to specific where in the word document we want the pasting of our Excel range to happen and how we want it to populate the page, either collapsing to the bottom / end of the page like its stated here or to have it justify another way. We also insert a a paragraph break in order to create separation from our body text and the incoming table.

`oWordRng.PasteAndFormat Worksheets("Excel2PP&OL").Range("C2:K23").CopyPicture`

Lastly, we will be accessing a specific range from our specified worksheet and will copy and paste it in one command to our outlook email item. Notice the specified info in the code, `Excel2PP&OL` and `C2:K23` will be replaced with your excel worksheet name and the specific range you want to copy on said worksheet.

 

## Creating the PowerPoint Slide 

Creating the PowerPoint slide with the table from Excel will follow a similar process to the steps we did for outlook.

 Opening PowerPoint and creating a new Presentation:

```
  Set PP = New PowerPoint.Application
  Set PPPres = PP.Presentations.Add

  PP.Visible = True
```

Specifying where to add the current slide and what kind of slide template we want to add. In this case we are adding a slide at position 1 which creates a slide that only has a tittle and the rest of it can be filled with our content. We then select the corresponding slide to be the active slide to be worked on.

```
  Set PPSlide = PPPres.Slides.Add(1, ppLayoutTitleOnly)

  PPSlide.Select
```

Similarly, to creating the Outlook email, we need to copy the specific range of cells from the correct worksheet. Some additional logic has been added to format the copied cells so that it looks better on our PowerPoint slide.

```
  Worksheets("Excel2PP&OL").Range("C2:K23").CopyPicture Appearance:=xlScreen, Format:=xlPicture
```

Now with the selected PowerPoint slide, we paste our copied cells and add a slight formatting to center the image on the center of our slide.

```
  PPSlide.Shapes.Paste.Select
  PP.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
```

Lastly, we add a title to our slide and do some simple formatting to clean up and center the title on the page.

 ```
  SlideTitle = "Test PowerPoint Slide" 
  PPSlide.Shapes.Title.TextFrame.TextRange.Text = SlideTitle
  PPSlide.Shapes.Title.TextEffect.Alignment = msoTextEffectAlignmentCentered
```

Finally, we activate our PowerPoint application to run all the following commands that we have specified.

`PP.Activate`

At the end of our script we reset all of our previously declared variables to “Nothing” in order to ensure that we clean up all the memory and avoid memory leaks in our app, this is a best practice when working with VBA.

```
  Set PPSlide = Nothing
  Set PPPres = Nothing
  Set PP = Nothing
  Set oLookApp = Nothing
  Set oLookItm = Nothing
  Set oLookIns = Nothing
  Set oWordDoc = Nothing
  Set oWordRng = Nothing
  Set oWordTable = Nothing
```

Now test your script by finding it in the Macros button under the developer tab and watch the magic happen!

# Digital Mock-Up in PowerPoint

In [this example](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/testfontsembedded.pptm) the menu [menuInfographics6.pptx](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInfographics6.pptx), which was converted from Java to MicroVBA ([macro.txt](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt)) and then imported in PowerPoint using the [MicroVBA interpreter](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/ReadMicroVBA.pptm), is used to simulate the menu behavior. This is ideal to be used as a **digital mock-up** before delivering the program. To know how this example was produced, please proceed to the section [Reading the Object](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/README.md#reading-the-object).

## Reading the Object

At the start, the file [testfontsembedded.pptm](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/testfontsembedded.pptm) was a blank presentation with a colored background that had just the macro **MicroVBA** and the file [macro.txt](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt) in the same directory. By runnning **MicroVBA**, that is, pressing **ALT+F8**, clicking **MicroVBA**, and clicking on _"Run"_, the file 
[macro.txt](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt) is read and interpreted. The result is that the menuInfographics6 appears in the blank slide. File [macro.txt](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt) actually has the instructions to create the menu _"menuInfographics6"_ using a subset of VBA language that is refered here as MicroVBA.

## Manipulating the Object

First, the whole object is selected and grouped. Then, clicking it with the right button and choosing **Format Shape**, the size is modified by clicking on **Size**, choosing **Lock aspect ratio**, and typing 80% in the field **Height** or **Width**. Now the object is moved to an appropriate place of the left side of the slide by grabbing it and pulling it by continuing to press the mouse button and moving the mouse. 

Once the object is placed in the desired location, it is ungrouped. Now each menu item is selected and grouped. The result is that the slide will have 5 shapes, each menu element corresponding to a shape in the correct order.

## Creating the Macro to Mock-Up the Menu

First, macros should be enabled. 

### Enabling Macros
This can be done in PowerPoint 2010/2013/2016/2019/365, by choosing **File**, clicking **PowerPoint options** at the bottom of the menu that appears:

![image](https://user-images.githubusercontent.com/80269251/117585769-2f718b80-b0e2-11eb-837c-e41290d45ffa.png)

Then clicking **Trust Center** on the left of the **PowerPoint Options** dialog box and clicking **Trust Center Settings** on right as indicated,

![image](https://user-images.githubusercontent.com/80269251/117586185-8d06d780-b0e4-11eb-8f97-387af9512fe0.png)

One clicks **Macro Settings** on the left of the dialog and finally choose **Enable all macros** (however, one should **never** execute a macro without examining it first).

![image](https://user-images.githubusercontent.com/80269251/117586010-99d6fb80-b0e3-11eb-8019-5ae875d7fd27.png)

### Starting the VBA editor
The VBA editor (also called the Integrated Development Environment, or simply IDE) is where one will work with VBA/macro code in PowerPoint. 

To start the VBA editor/IDE, one should press **ALT+F11**.

### Adding code
In the VBA editor, one should first make sure that the presentation (normally called **VBAProject (Presentation1)**) is highlighted in the left-hand pane.

Clicking **Insert**, then **Module** from the menu bar a new Module is created in the VBA IDE into the project. Modules are one of the several "containers" that can hold VBA code. The new module created is called **Module 2", since in module 

In the empty upper right window (Just below **(General)** and **Declarations**) the following is typed:

```VBA
Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long

Public Sub Wait(Seconds As Double)
    Dim endtime As Double
    endtime = DateTime.Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub

Sub clicked(item As String, text As String)
    Dim s As GroupShapes
    Set s = Application.ActivePresentation.Slides(1).Shapes(item).GroupItems
    Dim F1 As FillFormat, F2 As FillFormat
    Set F1 = s(2).Fill
    Set F2 = s(3).Fill
    Dim c1 As Long, c2 As Long
    c1 = F1.ForeColor
    c2 = F1.BackColor
    F1.ForeColor.RGB = c2
    F1.BackColor.RGB = c1
    F2.ForeColor.RGB = c1
    F2.BackColor.RGB = c2
    Dim t As Shape, x As Shape
    Set t = Application.ActivePresentation.Slides(1).Shapes(text)
    Dim I As Long
    For I = 1 To 5
        Set x = Application.ActivePresentation.Slides(1).Shapes("Text " & I)
        If Not x Is t Then x.Visible = False
    Next I
    Wait 0.01
    F1.ForeColor.RGB = c1
    F1.BackColor.RGB = c2
    F2.ForeColor.RGB = c2
    F2.BackColor.RGB = c1

    t.Visible = t.Visible Xor True
End Sub

Sub clicked1()
    clicked "Menu 1", "Text 1"
End Sub
Sub clicked2()
    clicked "Menu 2", "Text 2"
End Sub
Sub clicked3()
    clicked "Menu 3", "Text 3"
End Sub
Sub clicked4()
    clicked "Menu 4", "Text 4"
End Sub
Sub clicked5()
    clicked "Menu 5", "Text 5"
End Sub
```

#### External Function Calls
One can notice that an external function **WaitMessage** is used. When one is examining a macro one should pay special attention to external function calls like this. If a macro is copied from somewhere and function calls like this exist, it is a major red flag and one should check them carefully before running the macro. External function calls should be avoided at all times except when it is absolutely necessary. In this case, this function is called in order to release the application of being frozen while waiting the click on a menu item to be shown. If there is no waiting time the click flash is shown so quickly that it is not noticed. This problem comes from the fact that a click is processed only when the mouse button is released in VBA. If a click was processed when the mouse is clicked this problem would not exist and this function would not be necessary to call. However, to process a click on mouse release has one big advantage: the click flashing can be adjusted for a fixed amount of time as done here. The waiting time is managed by the function **Wait** that receives the number of seconds it should be waited.

#### Clicked Function

**Clicked** function is obviously the main function in this module. It is where the _"mock-up"_ really takes place. It receives the name of the item that has been clicked and the name of the object to be shown when this item is clicked. Assuming that each of the menu items with its assciated object (appearing when the item is clicked), both have already known names (according to this report this is not yet the case), this function does the following:

- Variable **"s"** receives the shape that corresponds to the name of the menu item clicked. As seen in [**Manipulating the Object**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/README.md#manipulating-the-object), this is actually a set of shapes that was grouped. The menu item is actually a set of different shapes. To simplify the identification of the menu item this grouping is essential and helps to deal with this item in a much higher level. Therefore, **"s"** actually receives the list of shapes in the group (which is kept in the **GroupItems** Collection).
- Variables **"F1"** and **"F2"** receive each one the **Fill** property of the second and third shape of the list. They correspond to both shapes that will change their aspect when the item is clicked.
- Variables **"c1"** and **"c2"** receive each one the **ForeColor** and the **BackColor** defining the gradient of the second shape. The third shape has the same colors but reversed.
- The colors of both gradients are reversed.
- Variable **"t"** receives the shape of the object that will appear when the item is clicked.
- A loop makes all other objects from other items invisible, except the one corresponding to the menu item that has been clicked.
- Waits one hundredth of a second.
- Puts the original gradient colors back in each of both shapes.
- Finally, it does an _"exclusive or"_ with the visibility status of the object associated with the item clicked. This has the effect of reversing the visibility status of this object. If the object was visible, it becomes invisible. If it was invisible, it becomes visible. This is to deal with case the item is clicked twice, making it appear and disappear while continuously clicking at the same item.

In other words, all this function is doing is reversing these gradients twice. The whole effect of it is a flashing sensation to give a feeback that the item has been clicked.

#### Functions Clicked1 to Clicked5

These functions just call the function **clicked** with the corresponding names of the items clicked and of their associated objects.

### Changing the Names of the Menu Items
As seen in [Clicked Function](https://github.com/nilostolte/MicroVBA-PowerPoint/tree/main/Example#clicked-function) and in [Functions Clicked1 to Clicked5](), the menu items should have names: "Menu 1", "Menu 2", "Menu 3", "Menu 4", and "Menu 5". This can be done in several different ways. One way is doing that with a scratch function, thus using VBA, typically as a macro language. Since this function is just used once it can be discarded after it is used. Another way is using the debugger and the **Immediate Window**. This approach is more cumbersome but it is also much more secure, because one can immediately see the result of the name changing by using **"Add Watch ..."**.

This approach is quite simple although also time consuming. To use the debugger one needs to set a breakpoint in the first statement of any function in the module. To do that one needs to click the first line of the function (one can pick **clicked1**, for example) and press **F9**. The line should appear marked with a brown line. 
Now one should click the Run button (a green right-facing arrowhead icon), choose **Run**, **Run Sub/User Form** from the menu bar or press **F5** to make the debugger stop at the line chosen. The line should now appear marked in yellow.

Now from the code in the function **clicked** just copy the expression **Application.ActivePresentation.Slides(1).Shapes**. Now right-clicking the function **clicked1** and clicking on **"Add Watch ..."**, paste the expression in the dialog field right below **Expression** and press **OK**. By expanding the item that appeared in the **Watches** window, one can see "Item 1", "Item 2", "Item 3", "Item 4", and "Item 5". By expanding "Item 1" and looking for the field "Name", one can see the default name PowerPoint has attributed to the first menu item.

Next, copying the expression **Application.ActivePresentation.Slides(1).Shapes** in the immediate window (to make it visible, just type **Ctrl-G**), and adding 
_**(1).Name = "Menu 1"**_ at the end of it one will obtain the following line:

**Application.ActivePresentation.Slides(1).Shapes(1).Name = "Menu 1"**

Then, just type **Enter** and the name in the Watches window will change to "Menu 1".

Now, conceal **Item 1** in **Watches** window, expand **Item 2**, and look for the **Name** field. After that just edit the content in the immediate window and change the **1** in **Shapes(1)** and **"Menu 1"** to **2**, thus obtining the line:

**Application.ActivePresentation.Slides(1).Shapes(2).Name = "Menu 2"**

Press **Enter** and this item name is also changed. Proceed in the same way to the other three items to change all the names of the menu items.

### Constructing a Transparent Mask over each Menu Item

This is a quite classical PowerPoint technique. Just create a "Rounded Rectangle" that tightly envelops each menu item.

Then just select each one at a time (since they are over the menu items one can select it by just clicking the corresponding menu item), select the **Insert** tab and click on **Action**. Selecting tab **Mouse Click** and clicking on **Run macro**, just choose **clicked1**, **clicked2**, **clicked3**, **clicked4**, or **clicked5** when refering to the Rounded Rectangle enveloping **Menu 1**, **Menu 2**, **Menu 3**, **Menu 4**, or **Menu 5**.

### Creating the Objects Associated with Menu Items

Now it is just a matter of creating all the objects that will appear when the menu items are clicked and renaming them as: **Text 1**, **Text 2**, **Text 3**, **Text 4**, and **Text 5** as shown above. The file is ready and should be exactly as shown in [testfontsembedded.pptm](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/testfontsembedded.pptm).



# MicroVBA-PowerPoint
MicroVBA is a VBA interpreter written in VBA to be used in **PowerPoint** in order to be able to import large vector graphics files. The advantages are: **vectorization of PowerPoint objects** (particularly vectorized texts), **high level solution to convert from other vector graphics formats**, portable way **to store vector graphics objects outside PowerPoint**, **smooth connectivity** with VBA programs already inside PowerPoint presentations, simple programmable solution for complex objects construction, **no limitations in the size of the files** and **more pertinent and helpful error messages**. It actually does not need full VBA conpatibility, since it can smoothly integrate with VBA programs in the Powerpoint presentation.

## MicroVBA Interpreter

The MicroVBA interpreter can be found inside the file [**ReadMicroVBA.pptm**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/ReadMicroVBA.pptm) under the name of **MicroVBA**. The macro **MicroVBA**, when executed, actually reads the file [**macro.txt**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt), which contains the MicroVBA program, and executes it. This program generates the vector objects shown in the presentation [menuInforgraphics6.pptx](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx).

The file [**macro.txt**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt) was generated automatically by a modified version of [FreeHep](https://github.com/nilostolte/FreeHep#freehep) project. This project allows to convert Java vector objects into MicroVBA
automatically. The Java program converted is [MenuInfographics6](https://github.com/nilostolte/Java-Vector-GUI/tree/main/MenuInfographics6), which is already available in this repository (please check its [README file](https://github.com/nilostolte/Java-Vector-GUI/tree/main/MenuInfographics6#menuinfographics6)).

A [Digital Mock-Up in PowerPoint](https://github.com/nilostolte/MicroVBA-PowerPoint/tree/main/Example) is also available. In this mock-up the menu used as an example ([MenuInfographics6](https://github.com/nilostolte/Java-Vector-GUI/tree/main/MenuInfographics6)) that was converted to MicroVBA (file [**macro.txt**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt)) is interpreted inside PowerPoint using [MicroVBA interpreter](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/ReadMicroVBA.pptm) 
and the resulting object is used in a [presentation](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/testfontsembedded.pptm) that mimicks the original [program](https://github.com/nilostolte/Java-Vector-GUI/tree/main/MenuInfographics6). When one clicks one item of the menu, a text appears explaining what the menu item activates when clicked. This example shows how our system can be used for automatic creation of digital mock-ups of Java interfaces in PowerPoint. 

A part of **macro.txt** is used in the section [Understanding PowerPoint Internal Path Representation](https://github.com/nilostolte/MicroVBA-PowerPoint#understanding-powerpoint-internal-path-representation) as an example to show how to create paths with subpaths using MicroVBA in PowerPoint.

## Overview
MicroVBA programs can theoretically run as macros in PowerPoint. However, due to limitations in the size of VBA procedures, 
VBA macros cannot be used as a general multi-purpose **Vector Graphics File Format**, since complex vector graphics files almost always have much more than 64kB. **MicroVBA solves this problem** and can definitely be used as a vector graphics file format with no limitations in size. Thus, **MicroVBA is a language that is between PostScript and PDF**, because it has limited programming features, even though it is more programmable than PDF and it can be connected to a full-fledged VBA program already in the PowerPoint presentation. It has all characteristics of PostScript with advantages, since Powerpoint is able to handle transparency (just like PDF). In this sense, MicroVBA uses PowerPoint as its rendering engine, in the same way PostScript files need an interpreter to display their content and PDF files also need a program (like Acrobat) to be rendered.

MicroVBA has three huge advantages over PostScript and PDF:
- It is VBA Basic, which is far much simpler than PostScript and PDF.
- It is a text file that can be part of a larger real program in VBA inside a PowerPoint presentation. That is, its excellent connectivity surpasses PostScript that can hardly be used in conjunction with something else except itself. It also largely surpasses PDF too, since PDF is a closed, mostly binary format, that also cannot be used outside of itself. In addition, the lightweight MicroVBA interpreter footprint does not overshadow simple VBA programs. PostScript interpreters and PDF renderers are just gigantic software.
- It can generate a very high end presentation using a tool that everybody knows how to use.

In addition to that, the choice of using real VBA commands makes MicroVBA code snippets easily debugged directly as PowerPoint macros, a  particularity that can only be compared with PostScript language. However, once an object is created in PowerPoint using a code snippet, it can then be copied, pasted, translated and scaled in a presentation. This principle is used below to explain how to construct PowerPoint paths. This is far superior and powerful than PostScript programming language because its code can only produce a rendering on the screen, except when it is explicitly converted to another format.

As one can easily see, MicroVBA opens the door to very powerful uses that were difficult to imagine without it. This is nearly an ideal solution because the heavy graphics can be left to MicroVBA to handle. The graphics can also come from other sources and be converted from other formats given by professional artists, for example.

What completes the use of MicroVBA are **convertions from other vector formats**. This has already been accomplished by a modification of FreeHEP project that generates a MicroVBA file from whatever a Java program shows (See the project [here](https://github.com/nilostolte/Java2PPT#java2ppt)). Even though this looks more like a programmer's solution, that is not really the case because the Java vector information can be stored in a file using an intermediate file format. This file can be read and the program can then convert the information into MicroVBA.

## Understanding PowerPoint Internal Path Representation

In this exposition the following snippet is used as a macro in the file [PPTPathAnalysis.pptm](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PPTPathAnalysis.pptm):

```vba
    Set MyPPT = Application.ActivePresentation
    Set MySlides = MyPPT.Slides
    Set MySlide = MySlides.Item(1)
    Set MyShapes = MySlide.Shapes
    Set MyPath = MyShapes.BuildFreeform(0, 286.75, 156.8356)
    MyPath.AddNodes 1, 1, 285.93304, 156.8356, 285.27, 156.17256, 285.27, 155.3556
    MyPath.AddNodes 1, 1, 285.27, 154.53864, 285.93304, 153.8756, 286.75, 153.8756
    MyPath.AddNodes 0, 0, 297.11002, 153.8756
    MyPath.AddNodes 1, 1, 297.92697, 153.8756, 298.59, 154.53864, 298.59, 155.3556
    MyPath.AddNodes 1, 1, 298.59, 156.17256, 297.92697, 156.8356, 297.11002, 156.8356
    MyPath.AddNodes 0, 0, 286.75, 156.8356
    MyPath.AddNodes 0, 0, 287.7641, 149.84082
    MyPath.AddNodes 1, 1, 286.9871, 149.58833, 286.56146, 148.75302, 286.81393, 147.97603
    MyPath.AddNodes 1, 1, 287.06644, 147.19902, 287.90173, 146.77338, 288.67874, 147.02557
    MyPath.AddNodes 0, 0, 298.5317, 150.22711
    MyPath.AddNodes 1, 1, 299.3087, 150.4796, 299.73434, 151.31491, 299.48184, 152.0919
    MyPath.AddNodes 1, 1, 299.22937, 152.86891, 298.39404, 153.29456, 297.61707, 153.04236
    MyPath.AddNodes 0, 0, 287.7641, 149.84082
    MyPath.AddNodes 0, 0, 286.75, 156.8356
    MyPath.AddNodes 0, 0, 290.89017, 143.5017
    MyPath.AddNodes 1, 1, 290.2292, 143.02158, 290.08267, 142.0954, 290.56277, 141.43443
    MyPath.AddNodes 1, 1, 291.04288, 140.77345, 291.9691, 140.62694, 292.63004, 141.10706
    MyPath.AddNodes 0, 0, 301.0113, 147.19637
    MyPath.AddNodes 1, 1, 301.67224, 147.67677, 301.81906, 148.60266, 301.33896, 149.26363
    MyPath.AddNodes 1, 1, 300.85855, 149.92459, 299.93237, 150.07141, 299.2717, 149.591
    MyPath.AddNodes 0, 0, 290.89017, 143.5017
    MyPath.AddNodes 0, 0, 286.75, 156.8356
    MyPath.AddNodes 0, 0, 295.8221, 138.4389
    MyPath.AddNodes 1, 1, 295.342, 137.77794, 295.48853, 136.85176, 296.14948, 136.37164
    MyPath.AddNodes 1, 1, 296.81046, 135.89153, 297.73663, 136.03806, 298.21674, 136.69902
    MyPath.AddNodes 0, 0, 304.30637, 145.08055
    MyPath.AddNodes 1, 1, 304.78647, 145.74152, 304.63965, 146.6674, 303.9787, 147.14781
    MyPath.AddNodes 1, 1, 303.31802, 147.62793, 302.3918, 147.48111, 301.9114, 146.82043
    MyPath.AddNodes 0, 0, 295.8221, 138.4389
    MyPath.AddNodes 0, 0, 286.75, 156.8356
    MyPath.AddNodes 0, 0, 302.07687, 135.14767
    MyPath.AddNodes 1, 1, 301.8244, 134.37068, 302.25003, 133.53537, 303.02704, 133.28288
    MyPath.AddNodes 1, 1, 303.80405, 133.0304, 304.63965, 133.45604, 304.89215, 134.23305
    MyPath.AddNodes 0, 0, 308.09338, 144.086
    MyPath.AddNodes 1, 1, 308.34586, 144.86299, 307.92023, 145.6983, 307.14322, 145.95079
    MyPath.AddNodes 1, 1, 306.3662, 146.20328, 305.5306, 145.77763, 305.2784, 145.00064
    MyPath.AddNodes 0, 0, 302.07687, 135.14767
    MyPath.AddNodes 0, 0, 286.75, 156.8356
    MyPath.AddNodes 0, 0, 282.31, 150.9156
    MyPath.AddNodes 0, 0, 279.35, 150.9156
    MyPath.AddNodes 0, 0, 279.35, 158.3156
    MyPath.AddNodes 1, 1, 279.35, 160.76648, 281.3391, 162.7556, 283.79, 162.7556
    MyPath.AddNodes 0, 0, 300.07, 162.7556
    MyPath.AddNodes 1, 1, 302.52087, 162.7556, 304.51, 160.76648, 304.51, 158.3156
    MyPath.AddNodes 0, 0, 304.51, 150.9156
    MyPath.AddNodes 0, 0, 301.55002, 150.9156
    MyPath.AddNodes 0, 0, 301.55002, 158.3156
    MyPath.AddNodes 1, 1, 301.55002, 159.13257, 300.88696, 159.79561, 300.07, 159.79561
    MyPath.AddNodes 0, 0, 283.79, 159.79561
    MyPath.AddNodes 1, 1, 282.97305, 159.79561, 282.31, 159.13257, 282.31, 158.3156
    MyPath.AddNodes 0, 0, 282.31, 150.9156
    Set MyShape = MyPath.ConvertToShape()
    Set MyFill = MyShape.Fill
    Set Color1 = MyFill.ForeColor
    Color1.RGB = 6573367
    Set MyLine = MyShape.Line
    MyLine.Visible = False
```

The first lines expose a major simplification used in MicroVBA: the parser cannot handle more than two indirections. This is not really a big problem, since indirections can be broken by subdividing them and storing them in a cascade of variables as shown above. Another limitation, is that expressions are not allowed. Also no VBA function is yet supported, except functions from objects as shown in the example. Functions can be easily wrapped into an object (named _"VBA"_, for example, where all the functions can be accessed as methods from this object). This has not been done because it is probable that this has already been done somewhere although it has not yet been found. This would be a useful contribution to this project. 

Finally, arithmetic operations can also be implemented as methods in the same object, for example, thus, avoiding complex expressions parsing which is not only slow but also increases too much the size of the interpreter. Complex expressions can be broken in simpler parts that are parsed and executed much faster than wasting time in the analysis of expressions. This principle is very similar to the one found in RISC processors. The reason RISC processors are so common nowadays is because it is better to have many more simple RISC cores than having a single complex one. It also consumes much less energy and has the potential to be much faster as all cores are used in a set of threads that are all executed in parallel. A similar reasoning applies here, since a simpler interpreter can revamp programs that already use the full VBA potential.

### Contructing Paths in PowerPoint

In the above example we can notice that paths are constructed using **BuildFreeform**. The parameters passed to the BuildFreeform are the coordinates of the first _"moveto"_ of the path (for a better understanding of paths please refer to [_"path definition"_](https://github.com/nilostolte/ClockWidget#paths) and [_"path commands"_](https://github.com/nilostolte/ClockWidget/blob/main/README.md#path-commands)). These coordinates are particularly important for complex paths in Powerpoint. Actually, there are no _"moveto"_ commands inside a 
path in Powerpoint. This is a huge limitation (also see [PowerPoint Caveats](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PowerPoint%20Caveats.md)) that is bypassed by an undocumented feature that can be used when defining complex paths. As in paths in other vector graphics languages, one can insert several subpaths inside a single path, but in Powerpoint this must be done according to certain rules. Although these rules are not always necessary they are simple to implement and understand:

- The first path in the group of subpaths must start with the initial _"moveto"_ declared in **BuildFreeform**
- The first path must be closed either using a _"lineto"_ to the coordinates indicated in **BuildFreeform** or they must be the last coordinates of a Bezier curve.
- The first _"moveto"_ of a subpath is actually a _"lineto"_ with its initial coordinates
- The subpath must be closed either with a "_lineto_" to the same coordinates or using these coordinates as the last coordinates of a Bezier curve.
- Then another _"lineto"_ to the initial coordinates in **BuildFreeform** must appear, except in the last subpath

In the example above we can identify 6 subpaths. The first obviously start with the coordinates declared in **BuildFreeform**:

``` vba
    Set MyPath = MyShapes.BuildFreeform(0, 286.75, 156.8356)
```

And is closed with the first _"lineto"_ back to the first coordinates of **BuildFreeform** 
```vba
    MyPath.AddNodes 0, 0, 286.75, 156.8356
```

Because the last coordinates of the last Bezier curve does not close the figure:

```vba
    MyPath.AddNodes 1, 1, 298.59, 156.17256, 297.92697, 156.8356, 297.11002, 156.8356
```

Now the following 5 subpaths start with _"lineto"_ and are closed with two _"lineto"_:

1. Second subpath
```vba
    MyPath.AddNodes 0, 0, 287.7641, 149.84082
          ⋮
    MyPath.AddNodes 0, 0, 287.7641, 149.84082
    MyPath.AddNodes 0, 0, 286.75, 156.8356
```
2. Third subpath
```vba
    MyPath.AddNodes 0, 0, 290.89017, 143.5017
          ⋮
    MyPath.AddNodes 0, 0, 290.89017, 143.5017
    MyPath.AddNodes 0, 0, 286.75, 156.8356
```
3. Fourth subpath
```vba
    MyPath.AddNodes 0, 0, 295.8221, 138.4389
          ⋮
    MyPath.AddNodes 0, 0, 295.8221, 138.4389
    MyPath.AddNodes 0, 0, 286.75, 156.8356
```
4. Fifth subpath
```vba
    MyPath.AddNodes 0, 0, 302.07687, 135.14767
          ⋮
    MyPath.AddNodes 0, 0, 302.07687, 135.14767
    MyPath.AddNodes 0, 0, 286.75, 156.8356
```
5. Sixth subpath
```vba
    MyPath.AddNodes 0, 0, 282.31, 150.9156
          ⋮
    MyPath.AddNodes 0, 0, 282.31, 150.9156
```
Notice that the last path does not need a _"lineto"_ to the begining of the first path. Also notice that _"lineto"_ commands are identified by four parameters. The first two should be 0, whereas the last two are the coordinates of the _"lineto"_. A _"curveto"_ command has eight parameters, where the first two are 1 and the other six parameters are the coordinates of the control points following the last point of the previous _"lineto"_ or _"curveto"_.

This solution has a drawback that can be used as a kind of WaterMark, as seen in [PowerPoint Caveats](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PowerPoint%20Caveats.md).

### Executing the Example

Opening the file [PPTPathAnalysis.pptm](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PPTPathAnalysis.pptm), one sees a blank presentation. One should make sure to [enable macros](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/Example/README.md#enabling-macros) in the PowerPoint application. The object must be created by the macro **example**. To see the macro, one should press **ALT+F8**. As a result one should get this window:

![image](https://user-images.githubusercontent.com/80269251/117228042-377db280-ade6-11eb-917a-66549aa73798.png)

#### Debugging the Example
Just click **Edit** instead of Run. The following window will open. Click on the statement **stop_here = 0** and press **F9**. This is a _"breakpoint"_, the point where the program will stop and wait until one allows it to continue.

![image](https://user-images.githubusercontent.com/80269251/117228832-dfe04680-ade7-11eb-9c23-01a129f24ad6.png)

#### Execution
Then click the Run button (a green right-facing arrowhead icon), choose **Run**, **Run Sub/User Form** from the menu bar or press **F5** to make the debugger stop at the line chosen, that is, to run the macro. One will have this content in the presentation, in the Powerpoint window:

![image](https://user-images.githubusercontent.com/80269251/117229502-25e9da00-ade9-11eb-9b8b-9857eead636e.png)

Going to this window, right click the object and choose **Edit Points**. One will get the following effect:

![image](https://user-images.githubusercontent.com/80269251/117229751-97298d00-ade9-11eb-96a6-09699409714c.png)

As one can see, the red lines are the _"lineto"_ commands to the first point of the path. These lines only appear when one choses **Edit Points** (they also appear if the PowerPoint file is converted to another vector graphics format, so this technique can also work as a sort of [watermark](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PowerPoint%20Caveats.md)). In fact all these elements are considered as a single object. One can also verify that in the VBA IDE window that the program stopped exactly where it was asked. One should leave it stopped at that point.

#### Invisible Points
In the Powerpoint window, the indicated points are a bit confusing because they do not correspond to the actual number of points created. When one clicks one of the points, other points appear. Some of these points really exist, and other are just graphics handles to change the geometry. This can be quite puzzling and confusing. 

To really see the actual points inside PowerPoint one has to use the debugger. Each point is stored in a structure called a **Node** and the path is stored in a Collection called **Nodes**. Since the debugger is stopped on a breakpoint, all the internal variables can be checked in the VBA IDE. To do that, one should first select the text **ActivePresentation.Slides(1).Shapes(1)**, right click the highlighted text, and click on **Add Watch...**, as indicated below:

![image](https://user-images.githubusercontent.com/80269251/117289921-30cd5a80-ae3b-11eb-8c6d-640d50a14578.png)

Then, in the **Add Watch** window just opened type **.Nodes** on the right of **ActivePresentation.Slides(1).Shapes(1)** as indicated below and click on the **OK** button.

![image](https://user-images.githubusercontent.com/80269251/117291954-9d495900-ae3d-11eb-8af2-1a085be2b56d.png)

The **Nodes** Collection will appear in the **Watches** subwindow. Expand **Watch** subwindow by press-clicking at the position indicated by the red cross inside the red circle and, still holding the mouse click button, pull it up as indicated below.

![image](https://user-images.githubusercontent.com/80269251/117293881-f619f100-ae3f-11eb-977b-d052f3340746.png)

Then click at the plus (**+**) sign on the left of **ActivePresentation.Slides(1).Shapes(1).Nodes** and on the right of the spectacles icon. To expand to see the complete name, just press-click at the vertical tab line on the left of **Value** and pull it to the right still holding the mouse click button.

Now click on the plus (**+**) sign on the left of **Item 1**, then on the left of **Points**, and **Points(1)** as shown below:

![image](https://user-images.githubusercontent.com/80269251/117295606-0d59de00-ae42-11eb-88cf-b359973111de.png)

#### The Pivot Point
One can see that the first point of the path had become **(289, 338)** because of the transformations applied in the original object after the macro **createobj** is called as indicated in the macro **example**. Now all the paths can be traced with these coordinates as pointed out in the **Notes** in the PowerPoint window. These notes were written to help to locate the different paths in the **Watch** subwindow as indicated above. The notes are shown here for convenience:

```
Node 1 – 14: element 1, starts at (289, 338) 
Node 15: line to (289, 338), closes element 1
Nodes 16 – 30: element 2
Node 31: line to (289, 338)
Nodes 32 – 46: element 3
Node 47: line to (289, 338)
Node 48 – 62: element 4
Node 63: line to (289, 338)
Nodes 64 – 78: element 5
Node 79:  line to (289, 338)
Nodes 80 – 100 : basket
```

A more detailed explanation about this object can be found in the section [The Resulting Stack Icon](https://github.com/nilostolte/MicroVBA-PowerPoint#the-resulting-stack-icon).

The first noticeable detail is that the first object is special as one can easily deduce from the notes. It has only 14 nodes, while the other "elements" have 15 nodes. This is because its closing _"lineto"_ command is mingled in node 15, which closes "element 1". Since the the first node can be thought as the first "moveto" of the whole path (passed as a parameter to **BuildFreeform**), this point becomes a kind of a pivot for all other subpaths as one could see when **Edit Points** has been chosen.

Afterwards each subpath has its own "moveto" (which is actually a "lineto") and a closing "lineto" (that could be last point of a curveto but not in this example). In addition to that, after each subpath, there is a "lineto" back to the point **(289, 338)**, as seen in nodes 31, 47, 63 and 79. The last subpath, identified as "basket" does not need to redirect to the point **(289, 338)** because there are no further subpaths following it. Therefore, the nodes 31, 47, 63 and 79 function as a kind of a marker to separate subpaths.

#### Discovering the method accidently
This method looks quite verbose but it allows putting any subpath anywhere in the path as demonstrated. This trick was actually discovered by accident, because texts in [**menugraphics6**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx) (marked in red below) which contains subpaths with "holes" (see section [Subtracting Shapes using Subpaths Winding](https://github.com/nilostolte/MicroVBA-PowerPoint#subtracting-shapes-using-subpaths-winding)) were the only texts to show up correctly. This was very puzzling since it is actually the most complex part of the menu items.

<p align="center">
  <img src="https://user-images.githubusercontent.com/80269251/117304424-ba852400-ae4b-11eb-9b57-9d404473bbd4.png">
</p>

Since that was highly unusual, and verifying what these texts were "doing correctly" that other texts and icons were obviously not (the errors are not shown above), the conclusion was this method, learned empirically, and shown above.

Microsoft apparently uses a different method to identify subpaths which is more compact but that is less flexible and more complex than the method shown above that looks quite robust.

As can be seen, the path used in this example is the icon of the second menu option above. The file  [**menugraphics6**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx) is given so the different objects can be examined in the way exposed above. The gear icon in the third menu option is another interesting example because it implements a hole in the gear. However, any text is also a good example since most of them contain letters that produces "holes" in a previous subpath.

Also, a very important point is that the file [**menugraphics6**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx) also demonstrates that vectorized texts are possible in PowerPoint presentations. Also, with vectorized texts, presentations can have a much better look, with a much better kerning and right justification.

## The Resulting Stack Icon

This example produces the icon of the second menu option of the file [**menugraphics6**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx) as shown above.

<p align="center">
  <img src="https://user-images.githubusercontent.com/80269251/117310644-74cb5a00-ae51-11eb-9aca-763522335f3d.png">
</p>

As explained above, the icon, which was produced by the macro **createobj**, was modified on the fly by the macro **example**. Since the object is indivisible one can scale it at will and all the points are scaled accordingly. Notice that the macro **createobj**, whose code is shown in the section [Understanding PowerPoint Internal Path Representation](https://github.com/nilostolte/MicroVBA-PowerPoint#understanding-powerpoint-internal-path-representation) above, is actually a snippet from a MicroVBA code ([**macro.txt**](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/MicroVBA%20Interpreter/macro.txt)) that produces the menu [menuinfographics6](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/menuInforgraphics6.pptx), supplied in this repository.

This icon represents a stack with five elements being pushed or popped from the stack. The stack itself is represented here as a "basket", as it was mentioned in section [The Pivot Point](https://github.com/nilostolte/MicroVBA-PowerPoint#the-pivot-point) above, while the five objects in the stack are simply referenced as "element 1" to "element&nbsp;5".

One can manually obtain paths with subpaths by using the **Combine Shapes** menu, which is hidden in PowerPoint 2010. It actually produces paths in a similar way that was demontrated above.

In this example the subpaths do not overlap and it is a perfect example of **Shape Union**. One can reproduce this by creating different objects, selecting all of them, and then applying **Shape Union** operation. However, how can one reproduce a **Shape Subtraction** operation using the scheme just presented above?

## Subtracting Shapes using Subpaths Winding

When applying **Shape Union** all shapes are assumed to be "winding" in the same direction. If the shapes are winding in opposite directions and if they overlap, a subtraction takes place as shown in the figure below:

<p align="center">
  <img src="https://user-images.githubusercontent.com/80269251/117321603-46527c80-ae5b-11eb-913b-4e3c90c0fbd6.png">
</p>
 
As one can see in the figure a small circle is subtracted from the big circle by winding the small circle in the opposite way. Also, the small circle should be subtracted from the big one. If the opposite is done both objects disappear. One can think in terms that a "negative object" does not exist. Thus, to obtain an actual subtraction, both objects must overlap and if one is inside the other, the small should be subratcted from the big one, otherwise both objects disappear.

To accomplish subtractions with subpaths created as shown in section [**Contructing Paths in PowerPoint**](https://github.com/nilostolte/MicroVBA-PowerPoint#contructing-paths-in-powerpoint) above, the object to be subtracted should appear first. Then, if one wants the following subpath to be subtracted from the previous one, they should wind their points in the opposite way from each other. By defining paths using MicroVBA as explained, subtractions are automatic when the subpaths are winding in opposite directions. These are conventions that are common in vector graphics. 
In general, one can modify the "winding rule" to stop behaving in this way to avoid subtractions in non-convex surfaces that intersects itself. In Java, this is done by using the **setWindingRule** function of a Path2D. Normally, the default winding rule is **WIND_NON_ZERO**, which allows operating in this way also in Java paths. However in PowerPoint it seems this rule is always **WIND_NON_ZERO** and it cannot aapparently be modified.

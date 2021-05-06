# MicroVBA-PowerPoint
MicroVBA is a VBA interpreter written in VBA to be used in **PowerPoint** in order to be able to import large vector graphics files. The advantages are: **vectorization of PowerPoint presentations** (no fonts needed), **fonts embedded in presentations**, can be used as a **Vector Graphics File Format** storage, simplification of complex objects construction, **no limitations in the size of the files** and **more pertinent and helpful error messages**. It can be extended to handle full VBA conpatibility. 

## Overview
MicroVBA programs can theoretically run as macros in PowerPoint. However, due to limitations in the size of VBA procedures, 
VBA macros cannot be used as a general multi-purpose **Vector Graphics File Format**, since complex vector graphics files almost always have much more than 64kB. **MicroVBA solves this problem** and can definitely be used as a programming language that can also function as a vector graphics format with no limitations in size. Thus, **MicroVBA is a language that is between PostScript and PDF**, because it has limited programming features, even though it is more programmable than PDF and it could be extended to a full-fledged VBA syntax. It has all characteristics of PostScript with advantages, since Powerpoint is able to handle transparency (as well as PDF). In this sense, MicroVBA uses PowerPoint as its rendering engine, in the same way PostScript files need an interpreter to display their content and PDF files also need a program (like Acrobat) to be rendered.

MicroVBA has three huge advantages over PostScript and PDF:
- It is VBA Basic, which is far much simpler than PostScript and PDF.
- It is a text file that can be part of a larger real program in VBA inside a PowerPoint file. That is, its excellent connectivity surpasses PostScript that can hardly be used in conjunction with something else except itself. It also largely surpasses PDF too, since PDF is a closed, mostly binary format, that also cannot be used outside of itself. In addition, the lightweight MicroVBA interpreter footprint does not overshadow simple VBA programs. PostScript interpreters and PDF renderers are just gigantic software.
- It can generate a very high end presentation using a tool that everybody knows how to use.

In addition to that, the choice of using real VBA commands makes MicroVBA code snippets easily debugged directly as PowerPoint macros, a  particularity that can only be compared with PostScript language. However, once an object is created in PowerPoint using a code snippet, it can then be copied, pasted, translated and scaled in a presentation. This principle is used below to explain how to construct PowerPoint paths. This is far superior and powerful than PostScript programming language because its code can only produce a rendering on the screen, except when it is explicitly converted to another format.

As one can easily see, MicroVBA opens the door to very powerful uses that were difficult to imagine without it.

What completes the use of MicroVBA are **convertions from other vector formats**. This has already been accoplished by a modification of FreeHEP project that generates a MicroVBA file from whatever a Java program shows. Even though this looks more like a programmer's solution, that is not really the case because the Java vector information can be stored in a file using an intermediate file format. This file can be read and the program can than convert the information into MicroVBA.

## Understanding PowerPoint Internal Path Representation.

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

The first lines expose a major simplification used in MicroVBA: the parser cannot handle more than two indirections. This is not really a big problem, since indirections can be broken by subdividing them and storing them in a cascade of variables as shown above. Another limitation, is that expressions are not allowed. Also no VBA function is yet supported, except functions from objects as shown in the example. Functions can be easily wrapped into an object (named _"VBA"_, for example, where all the functions can be accessed as methods from this object). This has not been done because it is probable that this has already been done somewhere although this has not yet been found. A useful contribution to this project would be this. 

Finally, arithmetic operations can also be implemented as methods in the same object, for example, thus, avoiding complex expressions parsing which is not only slow but also increases too much the size of the interpreter. Complex expressions can be broken in simpler parts that are parsed and executed much faster than wasting time in the analysis of expressions. This principle is very similar to the one found in RISC processors. The reason RISC processors are so common nowadays is because it is better to have many more simple RISC cores than having a single complex one. It also consumes much less energy and has the potential to be much faster as all cores are used in a set of threads that are all executed in parallel. A similar reasoning applies here, since a simpler interpreter can revamp programs that use already the full VBA potential.

### Contructing Paths in PowerPoint

In the above example we can notice that paths are constructed using **BuildFreeform**. The parameters passed to the BuildFreeform are the coordinates of the first _"moveto"_ of the path (for a better understanding of paths please refer to [_"path definition"_](https://github.com/nilostolte/ClockWidget#paths) and [_"path commands"_](https://github.com/nilostolte/ClockWidget/blob/main/README.md#path-commands)). These coordinates are particularly important for complex paths in Powerpoint. Actually, there are no _"moveto"_ commands inside a 
path in Powerpoint. This is a huge limitation that is bypassed by an undocumented feature that can be used when defining complex paths. As in paths in other vector graphics languages, one can insert several subpaths inside a single path, but in Powerpoint this must be done according to certain rules. Although these rules are not always necessary they are simple to implement and understand:

- The first path in the group of subpaths must start with the initial _"moveto"_ declared in **BuildFreeform**
- The first path must be closed either using a _"lineto"_ to the coordinates indicated in **BuildFreeform** or they must be the last coordinates of a bezier curve.
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

Because the last coordinates of the last bezier curve does not close the figure:

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
Notice that the last path does not need a _"lineto"_ to the begining of the first path. Also notice that _"lineto"_ commands are identified by four parameters. The first two should be 0, whereas the last two are the coordinates of the _"lineto"_. A _"curveto"_ command has eight parameters, where the first two are 1 and the other six parameters are the coordinates of the control points following the the last point of the previous _"lineto"_ or _"curveto"_.

### Executing the Example

Opening the file [PPTPathAnalysis.pptm](https://github.com/nilostolte/MicroVBA-PowerPoint/blob/main/PPTPathAnalysis.pptm), one sees a blank presentation. The object must be created by the macro **example**. To see the macro, one should click on tab **View** and click on **Macros**. As a result one should get this window:

![image](https://user-images.githubusercontent.com/80269251/117228042-377db280-ade6-11eb-917a-66549aa73798.png)

Just click **Edit** instead of Run. The following window will open. Make sure to click on the left side of the statement using the variable **stop_here**, indicated by the circle. This is a _"breakpoint"_, the point where the program will stop and wait until one allows it to continue.

![image](https://user-images.githubusercontent.com/80269251/117228832-dfe04680-ade7-11eb-9c23-01a129f24ad6.png)

Then click on the green triangle to play, that is, to run the macro. One will have this content in the presentation:

![image](https://user-images.githubusercontent.com/80269251/117229502-25e9da00-ade9-11eb-9b8b-9857eead636e.png)

Right click the object and choose **Edit Points**. One will get the following effect:

![image](https://user-images.githubusercontent.com/80269251/117229751-97298d00-ade9-11eb-96a6-09699409714c.png)

As one can see, the red lines are the _"lineto"_ commands to the first point of the path. These lines only appear when one chose **edit Points**. In fact all these elements are considered as a single object. Now, returning to the macro window, one can notice that the program stopped exactly where it was asked.

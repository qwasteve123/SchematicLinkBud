# SchematicLinkBud

This file is for design and management of link budget of Indoor distributed control system (DCS)  network. The aim for this project is to automate the tedious calculations in designing routing of feeder cable and choice of coupler in the whole rf network.

 ## Content

1. [Introduction](#md-header2-introduction)
2. [User Manual](#md-header2-user-manual)
    1. [Set Up Visio File](#md-header3-set-up)
    2. [(Optional) Set Up Shortcut Key for Macro](#md-header3-fast-key)
    3. [How To Use It](#md-header3-how-to-draw)
    4. [Edit Shape in Shape Data](#md-header3-edit-shape-data)
    5. [Fast Naming of Shapes](#md-header3-fast-naming)
    6. [Change Label Format](#md-header3-label-format)
    7. [Show or Export Link Budget](#md-header3-link-bud)
    8. [Find Shape Location in Visio](#md-header3-search-for-shape)
    9. [Change Shape of Connector (for layout drawing)](#md-header3-reroute)
    10. [Measure Length of Connector (for layout drawing)](#md-header3-to-length)

<h2 id="md-header2-introduction">Introduction </h2>
The main problem in indoor DCS system is the complexity in calculation to reach a balance of power output in every antenna. The system has a source to distrubute to every descending separation point (couplers) and eventually reach the output (antenna). The complexity of DCS system design comes from the limitation in construction site, mainly the pathway restricted and the antenna location pre-requested in tender by clients. To ensure the calculations reach the demanded power output (RSRP), link budget calculation is required. Without automation, 
we need to care about the choice of couplers, length and routing of feeder cables. Changing a coupler will alter the descendents' output which consisted in the same routing, which makes link budgets design difficult and complex.<br><br>
  
This project focus on calculation of RSRP output from antenna by subtracting the coupling loss and feeder loss in the whole pathway from source to ouput antenna. Visio is used for showcase of link budget and DCS network design. 

<h2 id="md-header2-user-manual"> User Manual </h2>

<h3 id="md-header3-set-up"> Set Up Visio File </h3>
If you already have a visio file with this code, just delete all contents in the document. The document is ready for use. <br><br>

1. For setting up visio document, you need to open Visio first. Press Development tab in toolbar and you can see Visual Basic Editor (VBE). Press it to enter VBE.
<br>___image required___
2. Download all files from this project, drag all VBA files (___.bas, .frm, .cls file only___) to VB Editor as follow:
<br>___image required___
3. Close VB Editor. Back to Visio file.
4. Choose ___File >>> Shapes >>> Open Stensils...,___ a file dialog will pop up. Choose ___.vss file___ in this project.
<br>___image required___
5. You can see the stensil with shapes shown up. It is recommended to put all shapes to document stensil. The document is set up.<br>
<br>___check if need to make dynamic connector___

<h3 id="md-header3-fast-key"> (Optional) Set Up Shortcut Key for Macro </h3>

1. Press run macro, from ___Macro in:___ dropdown list, select **A_main**
<br>___image required___
2. Select __ShowLinkBudget__ from list, Press __Option...__ to assign shortcut key. Press ___Shift + D___ and press __Ok__.
<br>___image required___
3. Select __ExportLinkBudget__ from list, Press __Option...__ to assign shortcut key. Press ___Shift + F___ and press __Ok__.
4. Back to __Macro in:__ dropdown list, select __DC_AutoNumbering__
<br>___image required___
5. Select __ActivateForm__ from list, Press __Option...__ to assign shortcut key. Press __Shift + W__ and press __Ok__. 

<h3 id="md-header3-how-to-draw"> How To Draw Schametic Diagram</h3>

1. The component shapes follow the conventional shape of couplers, splitter and antennae in schematic diagram. DCS system start from ___Start Block___ to ___Coupler/
2 way splitter/ 3 way splitter___ to ___Antenna___ and linked by ___Feeder cable___ in connection points. Connection method is based on schematic drawing.
2. To place shapes, press and drag shape from ___Document Stensil___ to your page.
<br>___image required___
3. To connect shapes, press ___Ctrl + 3___ to call dyamic connector and connect each shape.
 <br>___image required___
 
 <h3 id="md-header3-edit-shape-data"> Edit Shapes in Shape Data</h3>
 
1. To change the choice of couplers, right click on shape and click ___shape data___. Change coupler in ___coupling loss___ dropdown list.
<br>___image required___
2. To change the choice of feeder cable, right click on shape and click ___shape data___. Change feeder cable in ___feeder type___ dropdown list.
<br>___image required___
3. To change the antenna gain, right click on master shape in document stensil, click ___edit master___ . After the master shape is opened,right click on master shape and click ___shape data___. Change your antenna gain.
<br>___image required___
4. To change shape floor and item no. shown on shape text, edit __shape data___ window. Be careful that item no. can only accept ___integer___ while floor can be any text (string datatype). For more convinient method, read section [Fast Naming of Shapes](#md-header3-fast-naming)
<br>___image required___

 <h3 id="md-header3-fast-naming"> Fast Naming of Shapes</h3>
 
1. To change floor of multiple shapes, select multiple shapes, right click on shape and click ___shape data___ to change floor. Do not include shape that without floor input in your selection otherwise floor properties will not be shown in shape data window.
2. There are two methods to update item no. of shape:
     1. To name shapes item no. in consecutive order ( 1,2,3,...), Press ___Ctrl + Shift + W___ to active userform. (you can also run it in ___run       macro___). Go to tab ___Consecutive___, Press ___Start___ button to start naming. Click the shape you want to name, the number will increase in increment of 1.  
      To change the number added to shapes, enter new number in the textbox on left and click ___Change Number___ and continue clicking shapes.
      <br>___image required___
     2. Sometimes you need to update all shapes in increment of 1. In this situation, Press ___Ctrl + Shift + W___ to active userform, Go to tab ___Add number___,
      Enter the increment you want and press ___Selections Add___.
      <br>___image required___

<h3 id="md-header3-label-format"> Change Label Format</h3>

This project consists of the two label format, Normal label format and Lift label format. To switch between two label format, select multiple shapes, choose to ___Lift label format___ and click ___Change Label Format___.

<br>___image required___

 <h3 id="md-header3-link-bud"> Show or Export Link Budget</h3>
 
Press ___Ctrl + Shift + W___ to active, Go to tab ___Link bud___. Choose ___Show Link Budget___ or ___Export to excel___.
 
 <h3 id="md-header3-search-for-shape"> Find Shape Location in Visio </h3>
 
Press ___Ctrl + Shift + W___ to active userform, Go to tab ___Find Shape___. Select the ___Floor___, ___Component Type___. IF the component type is feeder cable, select ___Cable Type___. Click find. The target components are marked in red circles in visio.
 
<br>___image required___

<h3 id="md-header3-reroute"> Change Shape of Connector (for layout drawing) </h3>

1. Press ___Ctrl + Shift + W___ to active userform, Go to tab ___Route___. Select a connector in your page, click ___Reroute___ to start changing connector's shape. Click on page to change connector's shape.

<br>___image required___

2. To amend the connector shape, select a connector in your page, click ___Undo-reroute___ to return the orignial shape of connector (default routing by Visio).

<br>___image required___

3. The reroute function is used in layout drawing for adapt the curved wall or irregular shape in construction drawing and make layout drawing possible in Visio. Length measurement function is also equipped for direct measurement of feeder cable while drawing DCS system layout. To further automate the link budget design process, the layout will auto-generate and auto-update the schematic in the future.

<h3 id="md-header3-to-length"> Measure Length of Connector (for layout drawing) </h3>

 This function is to input length to connector with scale to construction layout. A few steps are required to set up page for finding sacle of visio page to construction layout:  
 1. __Right click___ any page tab, and click ___Insert page...___ to open Page Setup window.
 2. In Page properties, choose page type as ___Background___, select ___Millimeters___ in Measurement units dropdown list. A background page is inserted.
 3. Convert your construction layout from pdf/ dwg file to any picture files e.g. .jpg file. (higher quality recommended)
 4. Insert construction layout picture to background. Resize the picture relative to the background page. The background layout page is set up.
 
 <br>___image required___
 
To find the scale of Visio page to construction layout:
 1. Drag shape ___Layout Scale___ to your page. Align shape with any dimensions marked in the layout. (longer dimension is recommend to minimze measurement error)
 2.Go to Shape data of the shape, enter the dimension marked in the layout. Scale is set up.
 
 <br>___image required___
 
 To apply the actual length to connectors:
 1. Press ___Ctrl + Shift + W___ to active userform, Go to tab ___Route___. Select ___all___ connectors which you want to change its length, click ___Add Length___ button in userform. 
 
 <br>___image required___
 
 2. The ___Layout Scale___ only apply in one page. Each layout page requires to add a ___Layout Scale___.

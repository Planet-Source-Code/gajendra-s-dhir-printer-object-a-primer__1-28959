<div align="center">

## Printer Object \- A Primer


</div>

### Description

The focus of my article is to demystify the printer object and present it as a magnificient object, which can be used to churn out dashing  printouts without the support of any third party reporting tool.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gajendra S\. Dhir](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gajendra-s-dhir.md)
**Level**          |Intermediate
**User Rating**    |4.7 (89 globes from 19 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gajendra-s-dhir-printer-object-a-primer__1-28959/archive/master.zip)





### Source Code

<table width="100%" border=0><tr><td>
<div align="right">by,<br>
 <i><b>Gajendra S. Dhir</b></i><br>
 <font size="-1">Team Leader</font><br>
 <b>Data Spec</b><br>
 Bilaspur-CG, INDIA</div>
<p>All programmers creating software solutions for their client, invariably have to process data and generate output on paper, using the printer, in the form  of reports. There are many third party tools available in the market which are
 instrucmental in generating beautifully crafted reports. </p>
<p>I, too, have used such report writers, until recently, for my even my most
 simple printing requirements. That is until I discovered the power of the <code>printer</code>
 object. </p>
<p>Most literature on Visual Basic, including books and articles, generally explore
 this <code>printer</code> object superficially and this, I believe is, why most
 of us tend to overlook this simple yet powerful printing <i>tool</i>.</p>
<p>The focus of my article is to demystify the <code>printer</code> object and
 present it as a magnificient object, which can be used to churn out dashing
 printouts without the support of any third party reporting tool. For detailed
 syntaxes of the objects, statements, commands, properties and methods used here
 you are requested to refer to the excellent documentation provided by Microsoft.</p>
<p>The sub-topics covered in the article include...</p>
<ul>
 <li><a href="#selectprinter">Select Printer</a> </li>
 <li><a href="#pagesize">Set the Page dimensions</a></li>
 <li><a href="#newpage">Change to a new page</a></li>
 <li><a href="#enddoc">End of a Print Job</a></li>
 <li><a href="#killdoc">Cancel the Print Job</a></li>
 <li><a href="#headpos">Position the head</a></li>
 <li><a href="#printtext">Print the text</a></li>
 <li><a href="#justified">Justification - Left, Right, Center</a></li>
 <li><a href="#fontstyle">Font - Name, Size and Style</a></li>
 <li><a href="#printcolor">Print in Color</a></li>
 <li><a href="#directions">Points for Consideration</a></li>
</ul>
<h3></h3>
<h2><a name="selectprinter"></a>Select the printer </h2>
<p>Windows operating system allows you to install more than one printer. One of
 these is marked as the default printer and is offered as choice for printing
 by the applications. </p>
<p>VB provides us with the <code>Printers</code> collection and the <code>Printer</code>
 object to take care of our printing requirements.</p>
<p>The <code>printers</code> collection contains a list of the printers installed
 on your system. <code>Printers.Count</code> specifies the number of printers
 installed and any printer can be selected as <code>Printers(i)</code>, where
 <code>i</code> is a number between <code>0</code> and <code>Printer.Count-1</code>.</p>
<p>To get a list of all the printers installed we could use a code snipet, like
 this...</p>
<p><code>For i = 1 to Printers.Count - 1<br>
 &nbsp;&nbsp;&nbsp;&nbsp;Printer.Print Printers(i).Name<br>
 Next i<br>
 Printer.EndDoc</code></p>
<p>or </p>
<p><code>For Each P in Printers<br>
 &nbsp;&nbsp;&nbsp;&nbsp;Printer.Print P.Name<br>
 Next P<br>
 Printer.EndDoc</code></p>
<p>The <code>Printer</code> object represents the printer which has been marked
 as the default printer in the Windows environment.</p>
<p><i>The entire discussion here uses the <code>printer</code> object and can
 easily be modified to use the <code>Printers(i)</code> object.</i></p>
<h2><a name="pagesize"></a>Setup Page Dimensions</h2>
<p>The next thing that you must do is setup the dimensions of the paper on which
 you will be printing. Windows has 41 predefined paper sizes based on the standard
 paper sizes available around the world. Other than these if the size of the
 paper does not match any of these pre-defined sizes you may set it the custom
 size and specify your own height and width for the paper. The properties used
 here are <code>Printer.PaperSize</code>, <code>Printer.Height</code> and <code>Printer.Width</code>.</p>
<p> The more commonly used paper sizes are... </p>
<p><code>&nbsp;&nbsp;Printer.PaperSize = vbPRPSLetter<br></code>
 or<br>
<code>&nbsp;&nbsp;Printer.PaperSize = vbPRPSA4</code></p>
<p>Please refer to the Microsoft documentation for a complete list of paper size
 constants.</p>
<p>To use a custom size paper your code will look something like...</p>
<p><code>&nbsp;&nbsp;Printer.Height = 10 * 1440&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp'
 10 inch height x 1440 twips per inch<br>
 &nbsp;&nbsp;Printer.Width = 5 * 1440&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'&nbsp;&nbsp;5
 inch height x 1440 twips per inch</code></p>
<p>Any attempt to alter the height or the width of the <code>printer</code> object,
 automatically changes the <code>Printer.PaperSize</code> to <code>vbPRPSUser</code>.
</p>
<p>While you are at it, you may also want to setup the orientation of the paper.
</p>
<p><code>&nbsp;&nbsp;Printer.Orientation = vbPRORPortrait<br>
 </code> or<br>
 <code>&nbsp;&nbsp;Printer.Orientation = vbPRORLandscape</code></p>
<p>Any time during the print session you want to check the dimensions of the paper
 size you can refer to the <code>height</code> and <code>width</code> properties
 for the <code>printer</code> object.</p>
<p>While printing a page a typical use for the height is to compare the paper
 length with current position of the printer head and determine whether the next
 line can be printed on the same page or you should request for a new page.</p>
<p><i><b>Note</b>: Depending upon the printer driver installed for the printer
 it may or may not report an error is any of the printer properties is set beyond
 the acceptable range.</i></p>
<p></p>
<h2><a name="newpage"></a>Change to a new page</h2>
<p>Printing to the <code>printer</code> is done in page mode, i.e. the <code>printer</code>
 object sends data for printing to the operating system only after it is informed
 that the current page formatting is complete and is ready for printing. </p>
<p>In VB, this is accomplished by invoking the <code>NewPage</code> method like
 this... </p>
<p><code>Printer.NewPage</code></p>
<p>This method instructs the <code>printer</code> object to end the current page
 and advance to the next page. </p>
<h2><a name="enddoc"></a>End of Print Job</h2>
<p> When you have completed printing all the text and graphics that required to
 be printed in this print job the <code>printer</code> object must be so informed.
 You can do so using the <code>EndDoc</code> method.</p>
<p><code>Printer.EndDoc</code></p>
<p>This terminates a print operation and releases the document to the printer.
 If something has been printed on the current page it automatically issues a
 <code>Printer.NewPage</code> to complete printing of the page. If a <code>Printer.NewPage</code>
 has been issued just before the <code>Printer.EndDoc</code> method, no blank
 page is printed.</p>
<h2><a name="killdoc"></a>Cancel the Print Job</h2>
<p>There will be occasions when you may want to abort the print session. This
 may be in response to a cancel request from the user or any such situation requiring
 you to do so.</p>
<p>For such times we have been provided with the <code>KillDoc</code> method.</p>
<p><code>Printer.KillDoc</code></p>
<p>The difference of the <code>KillDoc</code> and the <code>EndDoc</code> methods
 is more apparent when the operating system's Print Manager is handling the print
 jobs. If the operating system's Print Manager is handling the print job <code>KillDoc</code>
 deletes the current print job and the printer receives nothing.</p>
<p>If Print Manager isn't handling the print job, some or all of the data may
 be sent to the printer before <code>KillDoc</code> can take effect. In this
 case, the printer driver resets the printer when possible and terminates the
 print job.</p>
<p></p>
<h2><a name="headpos"></a>Position the <i>Head</i></h2>
<p>We can get or set the position using the two properties, <code>Printer.CurrentX</code>
 and <code>Printer.CurrentY</code>. As obvious by their names the return the
 position on the X and Y axes respectively.</p>
<p><code>Label1.Caption = "(" & Printer.CurrentX & ", " & Printer.CurrentY & ")"</code>
</p>
<p>Alternately, you may use these very functions to position the printer head
 as per your requirement.</p>
<p><code>Printer.CurrentX = 1440<br>
 Printer.CurrentY = 1440</code></p>
<p>Remember 1 inch = 1440 twips. so this previous code snipet should position
 the printer head 1 inch from each the top and left margins. Similarly this next
 code snipet here will position the printer head at the center of the page (half
 of width and height).</p>
<p><code>Printer.CurrentX = Printer.Width / 2<br>
 Printer.CurrentY = Printer.Height / 2</code></p>
<p>Every print instruction issued to place text or graphic on the page moves the
 <code>CurrentX</code> and <code>CurrentY</code> and should be considered and,
 if necessary, taken care of before issuing the next print instruction.</p>
<h2><a name="printtext"></a>Print out the text</h2>
<p>To print use...<br>
 <br>
 <code>Printer.Print "Text to Print"</code> <br>
 <br>
 Printing starts at the location marked by the <code>CurrentX</code> and <code>CurrentY</code>.<br>
 <br>
 After the text as been printed the values of the <code>CurrentX</code> and <code>CurrentY</code>
 are changed to the new location. The new location is different when a , (comma)
 or a ; (semi-colon) is added at the end of the <code>Print</code> statement.
 Run the following code and compare the results...</p>
<b>Code 1</b>
<p><code>Printer.CurrentX = 0<br>
 Printer.CurrentY = 0<br>
 For i = 1 to 5<br>
 &nbsp;&nbsp;&nbsp;Printer.Print Printer.CurrentX &amp; &quot;, &quot; &amp;
 Printer.CurrentY<br>
 Next i</code></p>
<b>Code 2</b>
<p><code>Printer.CurrentX = 0<br>
 Printer.CurrentY = 0<br>
 For i = 1 to 5<br>
 &nbsp;&nbsp;&nbsp;Printer.Print Printer.CurrentX &amp; &quot;, &quot; &amp;
 Printer.CurrentY;<br>
 Next i</code></p>
<p>notice the ; (semi-colon) at the end of the print statement. </p>
and <b>Code 3</b>
<p><code>Printer.CurrentX = 0<br>
 Printer.CurrentY = 0<br>
 For i = 1 to 5<br>
 &nbsp;&nbsp;&nbsp;Printer.Print Printer.CurrentX &amp; &quot;, &quot; &amp;
 Printer.CurrentY,<br>
 Next i</code></p>
<p>in this case note the , (comma) at the end of the print statement.</p>
<h2><a name="justified"></a>Justification - Left, Right or Center</h2>
<p>Justification is accomplished with the help of two methods of the <code>printer</code>
 object, viz <code>Printer.TextHeight(Text)</code> and <code>Printer.TextWidth(Text)</code>,
 with which we can determine the about of vertical and horizontal space that
 will be occupied when you print the <code>Text</code>.</p>
<p>So in this example...</p>
<p><code>mTxt = "Gajendra S. Dhir"<br>
 TxtWidth = Printer.TextWidth(mTxt)</code></p>
<p><code>TxtWidth</code> is the amount of horizontal space required by the text
 in <code>mTxt</code> to print.</p>
<p>Let us see print this as Left, Right and Center Justified.</p>
<p><code>'to leave 1" Margins on the Left, Right and Top of the Printer<br>
 Printer.CurrentX = 1440<br>
 MaxWidth = Printer.Width - 1440*2<br>
 Printer.CurrentY = 1440<br>
 </code></p>
<p><i>Left Justified</i> is the simplest form of justification and the head position
 is already set.</p>
<p><code>Printer.Print mTxt</code></p>
<p>The printer head automatically moves to the starting point on the next line
 as there is no comma or semi-colon at the end of the <code>Print</code>. </p>
<p>Lets try <i>right justification</i>. We have <code>CurrentY</code> set for
 the next print statement. We need to set the <code>CurrentX</code>. Now we will
 require the <code>MaxWidth</code> and <code>TxtWidth</code> values, which we
 have ready with us (above).</p>
<p><code>' add 1440 is to maintain the 1&quot; Left Margin.<br>
 Printer.CurrentX = 1440 + (MaxWidth - TxtWidth)<br>
 Printer.Print mTxt</code></p>
<p>Similarly, you can achieve <i>center justification</i> </p>
<p> <code>Printer.CurrentX = 1440 + (MaxWidth - TxtWidth)/2&nbsp;&nbsp;&nbsp;&nbsp;'again
 1440 is to maintain Left Margin.<br>
 Printer.Print mTxt</code></p>
<p>This is all there is to printing text.</p>
<p>Ah yes ... just one more thing before we proceed. The above logic assume that
 <code>TxtWidth &lt; MaxWidth</code>. If the width of the text is greater than
 the maximum print width then you must separately process the text to either
 truncate it so that it fits the <code>MaxWidth</code> or split the lines suitably
 to simulate word-wrap.</p>
<p>For those interested, here's the entire code, </p>
<p><code> mTxt = "Gajendra S. Dhir"<br>
 TxtWidth = Printer.TextWidth(mTxt)<br>
 <br>
 </code><code>'to leave 1" Margins on the Top, Left and Right of the page<br>
 Printer.CurrentY = 1440<br>
 Printer.CurrentX = 1440<br>
 MaxWidth = Printer.Width - 1440*2<br>
 <br>
 'Left Justified - no extra work<br>
 Printer.Print mTxt<br>
 <br>
 'Right Justified<br>
 Printer.CurrentX = 1440 + (MaxWidth - TxtWidth)&nbsp;&nbsp;' add 1440 is to
 maintain the 1&quot; Left Margin<br>
 Printer.Print mTxt <br>
 <br>
 'Center Justified<br>
 Printer.CurrentX = 1440 + (MaxWidth - TxtWidth)/2&nbsp;&nbsp;&nbsp;&nbsp;'again
 1440 is to maintain Left Margin.<br>
 Printer.Print mTxt<br>
 <br>
 'Terminate Printing<br>
 Printer.EndDoc </code></p>
<h2><a name="fontstyle"></a>Font Name, Size and Style</h2>
<p>A wide variety of fonts, also known as typefaces, are available under the Windows
 operating system. Some are optimized for better screen appearance while others
 are designed with the printed output in mind. The printer that you use also
 has certain built-in fonts which you can access from your VB program.</p>
<p>The <code>Printer.FontCount</code> property tells you the number of fonts that
 are available in your system and are supported by current the printer. You can
 select the name of the font that you want to use for printing your text from
 the <code>Printer.Fonts</code> collection</p>
<p>To get a list of the names of the fonts available you can use a loop like this...</p>
<p><code>For i = 0 to Printer.FontCount-1<br>
 &nbsp;&nbsp;&nbsp;&nbsp;Printer.Print Printer.Fonts(i)<br>
 Next i</code> </p>
<p>or better still you could use the <code>Printer.Font.Name</code> property like
 this...</p>
<p><code>For i = 0 to Printer.FontCount-1<br>
 &nbsp;&nbsp;&nbsp;&nbsp;Printer.Font.Name = Printer.Fonts(i)<br>
 &nbsp;&nbsp;&nbsp;&nbsp;Printer.Print Printer.Font.Name<br>
 Next i</code> </p>
<p>to get a complete list of the fonts available with each <code>Font.Name</code>
 printed using that very typeface. </p>
<p>To determine or alter the size of the text that is being printed you must access
 the <code>Printer.Font.Size</code> property. Mayby something like this...</p>
<p><code>mSize = Printer.Font.Size<br>
 Printer.Font.Size = mSize + 4<br>
 Printer.Print &quot;THE TITLE TEXT&quot;<br>
 Printer.Font.Size = mSize</code></p>
<p>Other than this, control for <b>Bold</b>, <i>Italic</i>, <u>Underline</u> and
 <s>Strikethru</s> characteristics of a font that are available at your disposal
 as a Visual Basic programmer. These are boolean properties and take the values
 <code>True</code> or <code>False</code>. You may use these properties as...</p>
<p><code>&nbsp;Printer.Font.Bold = True </code>to enable and <code>False</code>
 to disable<br>
 <code>&nbsp;Printer.Font.Italic = True </code>to enable and <code>False</code>
 to disable<br>
 <code>&nbsp;Printer.Font.Strikethrough = True </code>to enable and <code>False</code>
 to disable<br>
 and<br>
 <code>&nbsp;Printer.Font.Underline = True </code>to enable and <code>False</code>
 to disable</p>
<p>The following code will give you a printout of all the printer fonts installed
 on your system along with the &quot;<b>bold</b>&quot; and &quot;<i>italic</i>&quot;
 texts printed next to the font name.</p>
<p><code>With Printer<br>
 &nbsp;&nbsp;For i = 0 to .FontCount-1<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Font.Name = Printer.Fonts(i)<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Print Printer.Font.Name;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'Note
 the ; (semi-colon) at the end of print<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Font.Bold = True<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Print &quot; Bold&quot;;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'Note
 the ; (semi-colon) at the end of print<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Font.Bold = False<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Font.Italic = True<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Print &quot; Italic&quot;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'Note
 <b>no</b> ; (semi-colon) at the end of print<br>
 &nbsp;&nbsp;&nbsp;&nbsp;.Font.Italic = False<br>
 &nbsp;&nbsp;&nbsp;&nbsp;If Printer.CurrentY + Printer.TextHeight(&quot;NextLine&quot;)
 &gt; Printer.Height - 720 Then<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Printer.NewPage<br>
 &nbsp;&nbsp;&nbsp;&nbsp;End If<br>
 &nbsp;&nbsp;Next i<br>
 End Width<br>
 <br>
 'Terminate Printing<br>
 Printer.EndDoc <br>
 </code></p>
<p>When working with the fonts you can also use <code>.FontName</code>, <code>.FontSize</code>,
 <code>.FontBold</code>, <code>.FontItalic</code>, <code>.FontStrikeThru</code>,
 <code>.FontUnderline</code> for <code>.Font.Name</code>, <code>.Font.Size</code>,
 <code>.Font.Bold</code>, <code>.Font.Italic</code>, <code>.Font.Strikethrough</code>,
 <code>.Font.Underline</code> used above.</p>
<h2><a name="printcolor"></a>Print in Color</h2>
<p>Printing in color adds to the presentation value of the final output. Let us
 add some color to our printing. </p>
<p>Use the <code>Printer.ColorMode</code> to enable or disable color printing for your color printer.</p>
<p><code>Printer.ColorMode = vbPRCMColor<br>
 </code> or<br>
 <code>Printer.ColorMode = vbPRCMMonochrome<br>
 </code></p>
<p>Depending on the printer installed, when you the set the printer to vbPRCMMonochrome
 prints in shades of black and white. </p>
<p>Once you have activated color printing you can control the color of the output
 through two properties two properties, <code>backcolor</code> and <code>forecolor</code>,
 of the <code>printer</code>, to control the color of the background and the
 foreground respectively. The color values can be assigned to these properties
 using the <code>RGB()</code> function.</p>
<p><code>Printer.ForeColor = RGB(255, 0, 0)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' For
 Text in Red Color<br>
 Printer.Print &quot;This text is in Red &quot;;<br>
 Printer.ForeColor = RGB(0, 0, 255)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' For Text in
 Blue<br>
 Printer.Print &quot;and this is in Blue&quot;<br>
 Printer.BackColor = RGB(255, 255, 0)&nbsp;&nbsp;&nbsp;' For Background in Yellow<br>
 Printer.Print &quot;The text here is Blue and the background is Yellow&quot;</code>
</p>
<p>Visual Basic has provided color constants for the standard colors, namely <code>vbBlue</code>,
 <code>vbRed</code>, <code>vbGreen</code>, <code>vbMagenta</code>, <code>vbCyan</code>,
 <code>vbYellow</code>, <code>vbBlack</code> and <code>vbWhite</code>.</p>
<h2>Points for Consideration</h2>
<p>Here are some tips which I think you will find useful during your exploration
 of the <code>printer</code> object...</p>
<ul>
 <li>You will need simple sub-routines to print text - left, right and center
  justified within a maximum width that you may specify. This will allow you
  to create the columns in a tabular report and adequately justify the text
  within the column.</li>
 <li>You could write a function to split long strings based on the print width
  to enable word wrapping. <font size="-1">See my previous code submitted titled
  <b>Split Strings for Word Wrapping</b>.</font></li>
 <li>The printer uses the same concept of device contexts that is used by Form
  and PictureBox Control. The difference is only in methods like <code>EndDoc</code>,
  <code>KillDoc</code>, <code>Cls</code> etc. Using code like...<br>
  <code>If Destination = "Printer" Then<br>
  &nbsp;&nbsp;&nbsp;&nbsp;Set objDC = Printer<br>
  Else<br>
  &nbsp;&nbsp;&nbsp;&nbsp;Set objDC = Picture1<br>
  Endif<br>
  objDC.Print "Hello! This is Gajendra"</code><br>
  you can easily create a print preview.</li>
</ul>
<p>I welcome and will appreciate constructive feedback and creative suggestions.</p></td></tr></table>


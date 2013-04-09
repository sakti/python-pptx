#####
Chart
#####

:Updated:  2013-04-07
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Introduction
============

One of the shapes available for placing on a PowerPoint slide is the *chart*.
The chart shape is perhaps the most complex of all the PresentationML shapes.
Contributing to its complexity is the fact that its source data is contained in
an embedded Excel spreadsheet.


Research protocol
=================

* ( ) Review MS API documentation
* ( ) Inspect minimal XML produced by PowerPoint® client
* ( ) Review and document relevant schema elements


MS API Analysis
===============

MS API method to add a chart is::

    Shapes.AddChart(Type, Left, Top, Width, Height)

There is a HasChart property on Shape to indicate the shape "has" a chart.
Seems like "is" a chart would be more apt, but I'm still looking :)

From the `Chart Members`_ page on MSDN.

Most interesting ``Chart`` members:

...


XML produced by PowerPoint® application
=======================================

Inspection Notes
----------------

...


XML produced by PowerPoint® client
----------------------------------

.. highlight:: xml

::

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <c:date1904 val="0"/>
      <c:lang val="en-US"/>
      <c:roundedCorners val="0"/>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
          <c14:style val="118"/>
        </mc:Choice>
        <mc:Fallback>
          <c:style val="18"/>
        </mc:Fallback>
      </mc:AlternateContent>
      <c:chart>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
          <c:layout/>
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$B$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 1</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$2:$B$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>4.3</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>2.5</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>3.5</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>4.5</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:ser>
              <c:idx val="1"/>
              <c:order val="1"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$C$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 2</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$C$2:$C$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>2.4</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>4.4</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>1.8</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>2.8</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:ser>
              <c:idx val="2"/>
              <c:order val="2"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$D$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 3</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$D$2:$D$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>2.0</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>2.0</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>3.0</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>5.0</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:dLbls>
              <c:showLegendKey val="0"/>
              <c:showVal val="0"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
            </c:dLbls>
            <c:gapWidth val="150"/>
            <c:axId val="2051737496"/>
            <c:axId val="2051748984"/>
          </c:barChart>
          <c:catAx>
            <c:axId val="2051737496"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="2051748984"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
            <c:noMultiLvlLbl val="0"/>
          </c:catAx>
          <c:valAx>
            <c:axId val="2051748984"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="l"/>
            <c:majorGridlines/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="2051737496"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="between"/>
          </c:valAx>
        </c:plotArea>
        <c:legend>
          <c:legendPos val="r"/>
          <c:layout/>
          <c:overlay val="0"/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
      </c:chart>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="1800"/>
          </a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
      <c:externalData r:id="rId1">
        <c:autoUpdate val="0"/>
      </c:externalData>
    </c:chartSpace>

Resources
=========

.. _Chart Members:
   http://msdn.microsoft.com/en-us/library/office/ff746468(v=office.14).aspx


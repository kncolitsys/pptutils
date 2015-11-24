pptutils
created by: todd sharp (todd@cfsilence.com)
version: 0.3
12/15/07

Version History:
0.1 - 12/07 - Initial Creation (not released)
0.2 - 12/07 - Rewrite component (not released)
0.3 - 12/07 - Initial release.  Upgrade POI library, add getPPTMetaData() and fixNull()
0.3.1 - 12/16/07 - Minor bug in samples.
0.3.2 12/16/07 - more minor bugs - one in pptutils.cfc and some in demos
0.4 - 12/19/07 - added note extraction to convertPowerPoint() method

About:
PPTUtils is a simple cfc that can be used to extract text and/or formatted markup (including images) from a PowerPoint file.  Since it is built upon HSLF it is subject to the limitations in that project:

"HSLF is the POI Project's pure Java implementation of the Powerpoint '97(-2007) file format. It does not support the new PowerPoint 2007 .pptx file format, which is not OLE2 based.

HSLF provides a way to read powerpoint presentations, and extract text from it. It also provides some (currently limited) edit capabilities."
 
To work with this CFC, just instantiate it like you would any component.

<!--- create the pptutils object --->
<cfset pptutils = createObject("component", "pptutils.com.pptutils").init() />

To extract text call the extractText() method.
To create an array of slides (which will include data about formatting/positioning, etc) call the convertPowerPoint() method.
To extract metadata about the PPT file call the getPPTMetaData() method.

For sample implementations see the samples folder in this package.

Also see docs/pptutils_api.html

If you should find this component usefull please consider visiting my wishlist (http://www.amazon.com/gp/registry/wishlist/2PTWNTIRNTIKS/)

Credits:
pptutils uses JavaLoader by Mark Mandel (http://javaloader.riaforge.org/)
pptutils utilizes POI/HSLF which is part of the Apache Project (http://poi.apache.org/)
Licenses applicable to the respective owners.


License:
pptutils:
Copyright 2006 Todd Sharp
  
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
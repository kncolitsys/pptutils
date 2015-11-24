
<!--- 
License:
pptutils:
Copyright 2007 Todd Sharp
  
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 --->

<!--- get the path to the ppt --->
<cfset pptFile = expandPath("verity.ppt") />
<!--- create the pptutils object --->
<cfset pptutils = createObject("component", "pptutils.com.pptutils").init() />
<!--- get the ppt in html format --->
<cfset ppt = pptutils.readPowerPoint(pathToPPT=pptFile) />
<html>
<head>
</head>
<body>
<!--- 
<cfdump var="#ppt#">
<cfabort>  
 --->

<cfset imgStruct = structNew() />

<cfoutput>
<cfloop from="1" to="#arrayLen(ppt)#" index="i">
	<cfset slide = ppt[i] />
	<cfset imgName = "" />
	
	<cfif len(slide.slideBackgroundImage.imgData) and listFindNoCase(getWriteableImageFormats(), slide.slideBackgroundImage.imgType)>
		<cfset imgAsBase64 = toBase64(slide.slideBackgroundImage.imgData) />
		<cfset findImgArr = structFindValue(imgStruct, imgAsBase64) />
		
		<cfif not arrayLen(findImgArr)>
			<cfset imgName = "bg_" & structCount(imgStruct) + 1 & "." & slide.slideBackgroundImage.imgType />
			<cfset structInsert(imgStruct, imgName, imgAsBase64)>
			<cfset bgImg = imageNew(slide.slideBackgroundImage.imgData) />
			<cfset ImageScaleToFit(bgImg, slide.slideWidth, slide.slideHeight) />
			<cfset imgDest = expandPath(imgName) />
			<cfimage action="write" source="#bgImg#" destination="#imgDest#" overwrite="true" />
		<cfelse>
			<cfset imgName = findImgArr[1].key />
		</cfif>
	</cfif>
	
	<div style="position:relative; height:#slide.slideHeight#px;width:#slide.slideWidth#px;background: url(#imgName#) no-repeat;"><!--- this div represents a slide --->
		<!--- textboxes --->
		<cfloop from="1" to="#arrayLen(slide.textBoxes)#" index="j">
			<cfset tBox = slide.textBoxes[j] />
			<div style="position:absolute; left:#tBox.x#px; top:#tBox.y#px; height:#tBox.height#px; width:#tBox.width#px;">
				<cfloop from="1" to="#arrayLen(tBox.textSpans)#" index="k">
					<cfset tSpan = tBox.textSpans[k] />
					<div style="color:rgb(#tSpan.fontColor#); font-family:#tSpan.fontFamily#; font-size:#tSpan.fontSize#; font-weight:#tSpan.fontWeight#; text-align:#tSpan.textAlign#; text-decoration:#tSpan.textDecoration#;">
						#repeatString(repeatString("&nbsp;", 3), tSpan.indentLevel)#<!--- <cfif len(tSpan.bulletChar)>#chr(tSpan.bulletChar)#</cfif> --->
						#tSpan.text#
					</div>
				</cfloop>
			</div>
		</cfloop>
		<!--- images --->
		<cfloop from="1" to="#arrayLen(slide.images)#" index="k">
			<cfset img = slide.images[k] />
			<cfif listFindNoCase(getReadableImageFormats(), img.imgType)>
				<cfset i = imageNew(img.imgdata) />
				<cfset imageResize(i,img.width, img.height )>
				<div style="position:absolute; left:#img.x#px; top:#img.y#px;">
					<cfimage action="writeToBrowser" source="#i#">
				</div>
			</cfif>
		</cfloop>
	</div><!--- end slide --->
</cfloop>
</cfoutput>

</body>
</html>

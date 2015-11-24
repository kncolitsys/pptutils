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
<cfset ppt = pptutils.extractText(pathToPPT=pptFile) />

<cfdump var="#ppt#">
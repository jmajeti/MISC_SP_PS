/* 
/*  SharePoint JavaScript based cascaded dropdowns
/*  Published on spcd.codeplex.com
/*  Visit www.sharepointboris.net for more SharePoint tips and Tricks
/*  Tested on SharePoint 3.0 / 2007
*/ 

var setupComplete = false;
var wssSvcUrl = L_Menu_BaseUrl + "/_vti_bin/lists.asmx";
//Initiate function
function addHandler() {
	for (i = 0; i < CascadingDropdowns.length; i++) {
		if(CascadingDropdowns[i].parentLookup.isDropDown) {
			CascadingDropdowns[i].parentLookup.Object.onchange = function() { FilterChooicesForMyChild(this); }
		}
		else {
			CascadingDropdowns[i].parentLookup.Opthid.onpropertychange = function() { FilterChooicesForMyChild(this); }
		}
	}
	setupComplete = true;
}

//support functions
function getField(fieldType,fieldTitle) {   
     var docTags = document.getElementsByTagName(fieldType);   
     for (var i=0; i < docTags.length; i++) {   
         if (docTags[i].title == fieldTitle) {   
             return docTags[i];   
         }   
     }   
     return false;   
}  

//Object for working with web services
function WssSvcCall() {
	this.soapQuery = "";
	this.url = "";
	this.returnFunctionName = function() { return; };
	
	this.Submit = function() {
      http_request = false;
      if (window.XMLHttpRequest) { // Mozilla, Safari,...
         http_request = new XMLHttpRequest();
         if (http_request.overrideMimeType) {
         	// set type accordingly to anticipated content type
            //http_request.overrideMimeType('text/xml');
            http_request.overrideMimeType('text/html');
         }
      } else if (window.ActiveXObject) { // IE
         try {
            http_request = new ActiveXObject("Msxml2.XMLHTTP");
         } catch (e) {
            try {
               http_request = new ActiveXObject("Microsoft.XMLHTTP");
            } catch (e) {}
         }
      }
      if (!http_request) {
         alert('Cannot create XMLHTTP instance');
         return false;
      }
      
      http_request.onreadystatechange = this.returnFunctionName;
      http_request.open('POST', this.url, true);
      http_request.setRequestHeader("Content-type", "application/soap+xml");
      http_request.setRequestHeader("Content-length", this.soapQuery.length);
      http_request.send(this.soapQuery);
   }
}

//Object to hold the lookup field and needed properties
function LookupField(LookupFieldTitle) {
	this.Object = false;
	this.Opthid = false;
	this.isDropDown = true;
	
    if(getField('select',LookupFieldTitle))
    {
        //if lookup has 19 or less items - SELECT
        this.Object = getField('select',LookupFieldTitle);
    }
    else
    {
        //if it has 20 or more items - INPUT
        this.Object = getField('input',LookupFieldTitle);
        this.Opthid = document.getElementById(this.Object.optHid);
		this.isDropDown = false;
    }
}

var CascadingDropdowns = new Array();
//Object to hold cascading relationship info
function cascadeDropdowns(ParentDropDownTitle, ChildDropDownTitle, Child2ParentFieldIntName, ChildListNameOrGuid, ChildLookupTargetField) {
	this.parentLookup = new LookupField(ParentDropDownTitle);
	this.childLookup = new LookupField(ChildDropDownTitle);
	this.childList = ChildListNameOrGuid;
	this.child2ParentLink = Child2ParentFieldIntName;
	this.childLookupTargetField = ChildLookupTargetField;

	CascadingDropdowns.push(this);
}

function FilterChooicesForMyChild(triggerObject) {
	if(!setupComplete) return; //security in IE not to trigger filter on load
	for(i = 0; i < CascadingDropdowns.length; i++) {
		if(CascadingDropdowns[i].parentLookup.Object == triggerObject || CascadingDropdowns[i].parentLookup.Opthid == triggerObject) {
			var CascadingDropdown = CascadingDropdowns[i];
			var wssSvc = new WssSvcCall();
			wssSvc.soapQuery = getQuery2Run(CascadingDropdown);
			wssSvc.url = wssSvcUrl;
			wssSvc.returnFunctionName = function() { filterChildLookup(CascadingDropdown); }
			wssSvc.Submit();
		}
	}
}

function getQuery2Run(CascadingDropdown) {
	var selectedId;
	if(CascadingDropdown.parentLookup.isDropDown) {
		selectedId = CascadingDropdown.parentLookup.Object.options[CascadingDropdown.parentLookup.Object.selectedIndex].value;
	}
	else {
		selectedId = CascadingDropdown.parentLookup.Opthid.value;
	}
	var result = '<?xml version="1.0" encoding="utf-8"?>' +
	'<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">' +
	  '<soap12:Body>' +
		'<GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
		'<listName>' + CascadingDropdown.childList + '</listName>' +
			'<query>' +
				'<Query>' +
				(selectedId==0?'':'<Where><Eq><FieldRef LookupId="TRUE" Name="' + CascadingDropdown.child2ParentLink + '" /><Value Type="Counter">' + selectedId + '</Value></Eq></Where>') + 
				'<OrderBy><FieldRef Name="' + CascadingDropdown.childLookupTargetField + '" /></OrderBy>' +
				'</Query>' +
			'</query>' +
			'<viewFields><ViewFields>' +
				'<FieldRef Name="' + CascadingDropdown.childLookupTargetField + '" /><FieldRef Name="ID" />' + 
			'</ViewFields></viewFields>' +
		'</GetListItems>' +
	  '</soap12:Body>' +
	'</soap12:Envelope>'
	return result;
}

function filterChildLookup(CascadingDropdown) {
  if (http_request.readyState == 4) {
	 if (http_request.status == 200) {
		var xmlResult = parseXML(http_request.responseText);
		var resultNodes = xmlResult.getElementsByTagName('z:row');
		if(CascadingDropdown.childLookup.isDropDown) {
			var startAt = 0;
			if(CascadingDropdown.childLookup.Object.options[0].value == 0) startAt = 1;
			CascadingDropdown.childLookup.Object.options.length = startAt;
			for(y = 0; y < resultNodes.length; y++) {
				CascadingDropdown.childLookup.Object.options[y + startAt] = new Option(attributeValue(resultNodes[y], "ows_" + CascadingDropdown.childLookupTargetField), attributeValue(resultNodes[y], "ows_ID"), false, false);
				if(CascadingDropdown.childLookup.Object.options.length > 0) CascadingDropdown.childLookup.Object.options[0].selected = "selected";
			}
		}
		else {
			var choices = CascadingDropdown.childLookup.Object.choices;
			if(choices.substr(choices.indexOf("|"),3) == "|0|") choices = choices.substr(0, choices.indexOf("|")+2);
			var choicesArr = new Array();
			for(y = 0; y < resultNodes.length; y++) {
				choicesArr.push(attributeValue(resultNodes[y], "ows_" + CascadingDropdown.childLookupTargetField) + "|" + attributeValue(resultNodes[y], "ows_ID"));
			}
			choices += (choicesArr.length==0?"":"|") + choicesArr.join('|');
			CascadingDropdown.childLookup.Object.choices = choices;
			CascadingDropdown.childLookup.Object.value = "";
			CascadingDropdown.childLookup.Opthid.value = "";
		}
	 } else {
		alert('There was a problem with the request.');
	 }
  }
}

//xml parsing support functions
function parseXML(inputString) {
	if (window.DOMParser) {
		parser=new DOMParser();
		xmlDoc=parser.parseFromString(inputString,"text/xml");
	}
	else { // Internet Explorer
		xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
		xmlDoc.async="false";
		xmlDoc.loadXML(inputString); 
	}
	return xmlDoc;
}

function attributeValue(node, attributeName) {
	var attributesCollection = node.attributes;
	for(atv = 0; atv < attributesCollection.length; atv++) {
		if(attributesCollection[atv].name == attributeName) return attributesCollection[atv].value
	}
	return "";
}

//Start the party after the DOM is loaded
_spBodyOnLoadFunctionNames.push('addHandler');

Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 

URLColeccion="http://..."

'Recupera los usuarios que son grupos de dominio y los guarda en un array
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
		 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
		 "<soap:Body>"+_
		 "<GetUserCollectionFromSite xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/' />"+_
		 "</soap:Body>"+_
		 "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/directory/GetUserCollectionFromSite"
URLServicio = URLColeccion+"/_vti_bin/UserGroup.asmx"

EjecutaPeticion

ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Usuarios = ObjetoDOM.getElementsByTagName("User") 
Dim GruposdeDominio
ReDim GruposdeDominio(-1)
For i=0 To Usuarios.length-1
	'Filtra los grupos de dominio
	If Usuarios.item(i).getAttribute("IsDomainGroup")="True" Then
		ReDim Preserve GruposdeDominio(UBound(GruposdeDominio) + 1)
		GruposdeDominio(UBound(GruposdeDominio)) = Usuarios.item(i).getAttribute("LoginName")
	End If
Next

'Recupera los roles y los guarda en array
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<GetSite xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
         "</soap:Body>"+_
         "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetWeb"
URLServicio = URLColeccion+"/_vti_bin/SiteData.asmx"

EjecutaPeticion

ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set NodoRoles= ObjetoDOM.selectSingleNode("//Permissions")
ObjetoDOM.loadXML(NodoRoles.text)
Set Roles = ObjetoDOM.getElementsByTagName("Permission") 

'Roles Definidos.
'Es posible que existan roles diferentes con la misma mascara
For i=0 To Roles.length-1
Mascara=Mascara+Roles.item(i).getAttribute("Mask")+VBTab
NombreRol=NombreRol+Roles.item(i).getAttribute("RoleName")+VBTab
MascaraBinario=MascaraBinario+DecimalBinario(Roles.item(i).getAttribute("Mask"))+VBTab
Next

'Roles Combinados. La UNION de dos roles puede generar un rol diferente
For x=0 To Roles.length-1
	For y=x+1 To Roles.length-1
		Mx=CLng(Roles.item(x).getAttribute("Mask"))
		My=CLng(Roles.item(y).getAttribute("Mask"))
		Mxory =(Mx or My)
		If Mxory <> Mx And Mxory <> My Then
			Nivelxy=Roles.item(x).getAttribute("RoleName")+", "+Roles.item(y).getAttribute("RoleName")
			Mascara=Mascara&Mxory&VBTab
			NombreRol=NombreRol+Nivelxy+VBTab
			MascaraBinario=MascaraBinario&DecimalBinario(Mxory)&VBTab
		End If
	Next
Next 

Mascara=Split(Mascara,VBTab)
NombreRol=Split(NombreRol,VBTab)
MascaraBinario=Split(MascaraBinario,VBTab)
	
'Peticion del nombre de cada uno de los Sitios de la coleccion
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<GetSite xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
         "</soap:Body>"+_
         "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetSite"
URLServicio = URLColeccion+"/_vti_bin/SiteData.asmx"

EjecutaPeticion

Set Fichero = ObjetoArchivo.CreateTextFile("SalidaPermisosColecciondeSitios.txt", True) 
Fichero.WriteLine ("Sitio"+VBTab+"Lista"+VBTab+"Tipo [Plantilla]"+VBTab+"Hereda permisos"+VBTab+"Nombre"+VBTab+"Tipo"+VBTab+"Permisos"+VBTab+"Niveles")

'Recupera el nombre cada subsitio
ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Sitios = ObjetoDOM.getElementsByTagName("_sWebWithTime")

For i=0 To Sitios.length-1
	Set SitioURL = Sitios.item(i).selectSingleNode("Url")
	'Herencia en sitios
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
			 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
			 "<soap:Body>"+_
			 "<GetWeb xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
			 "</soap:Body>"+_
			 "</soap:Envelope>"
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetWeb"
	URLServicio = SitioURL.text+"/_vti_bin/SiteData.asmx"
	EjecutaPeticion
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set Herencia= ObjetoDOM.selectSingleNode("//InheritedSecurity")
	Fichero.WriteLine (SitioURL.text+VBTab+VBTab+VBTab+Herencia.text)
	If Herencia.text="false" Then
		'Obtiene los Permisos del sitio (Sólo cuando la herencia está rota)
		Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
				 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
				 "<soap:Body>"+_
				 "<GetPermissionCollection xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/'>"+_
				 "<objectName></objectName>"+_
				 "<objectType>Web</objectType>"+_
				 "</GetPermissionCollection>"+_
				 "</soap:Body>"+_
				 "</soap:Envelope>"
		LeePermisos
	End If
	
	'Petición de la coleccion de listas de cada sitio
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
             "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
             "<soap:Body>"+_
			 "<GetListCollection xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
			 "</soap:Body>"+_
			 "</soap:Envelope>"
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetListCollection"
	URLServicio = SitioURL.text+"/_vti_bin/SiteData.asmx"
	
	EjecutaPeticion
	
	'Recupera datos de cada una de las listas
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set Listas = ObjetoDOM.getElementsByTagName("_sList")
	
	For j=0 To Listas.length-1
		Set ListaTitulo = Listas.item(j).selectSingleNode("Title")
		Set ListaTipo = Listas.item(j).selectSingleNode("BaseType")
		Set ListaPlantilla = Listas.item(j).selectSingleNode("BaseTemplate")
		Set ListaHerencia = Listas.item(j).selectSingleNode("InheritedSecurity")
		Fichero.WriteLine (VBTab+ListaTitulo.text+VBTab+ListaTipo.text+" ["+ListaPlantilla.text+"]"+VBTab+ListaHerencia.text)
		If ListaHerencia.text="false" Then
			'Obtiene los Permisos de la lista (Sólo cuando la herencia está rota)
			Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
					 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
					 "<soap:Body>"+_
					 "<GetPermissionCollection xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/'>"+_
					 "<objectName>"+ListaTitulo.text+"</objectName>"+_
					 "<objectType>List</objectType>"+_
					 "</GetPermissionCollection>"+_
					 "</soap:Body>"+_
					 "</soap:Envelope>"
			LeePermisos
		End If
	Next
Next 

Fichero.Close

WScript.Echo "Informe obtenido"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Sub LeePermisos
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/directory/GetPermissionCollection"
URLServicio = SitioURL.text+"/_vti_bin/Permissions.asmx"
EjecutaPeticion
ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Permisos = ObjetoDOM.getElementsByTagName("Permission")
For k=0 To Permisos.length-1
	If Permisos.item(k).getAttribute("MemberIsUser")="True" Then
		Nombre=Permisos.item(k).getAttribute("UserLogin")
		'Determina si es Usuario o Grupo de Dominio
		Tipo="Usuario"
		For l=0 To UBound(GruposdeDominio)
			If Nombre=GruposdeDominio(l) Then
				Tipo="Grupo de Dominio"
			End If
		Next
	Else
		Nombre=Permisos.item(k).getAttribute("GroupName")
		Tipo="Grupo SharePoint"
	End If
	RolNivelBinario=RolBinario(Permisos.item(k).getAttribute("Mask"))
	RolNivel = Rol(Permisos.item(k).getAttribute("Mask"))
	Fichero.WriteLine (VBTab+VBTab+VBTab+VBTab+Nombre+VBTab+Tipo+VBTab+RolNivelBinario+VBTab+RolNivel)
Next
End Sub

Private Function Rol (Texto)
Rol="Desconocido"
For z=0 To UBound(Mascara)-1
	If Mascara(z)=Texto Then
		Rol=NombreRol(z)
		Exit For
	End If
Next
End Function

Private Function RolBinario (Texto)
RolBinario=Texto
For z=0 To UBound(Mascara)-1
	If Mascara(z)=Texto Then
		RolBinario=MascaraBinario(z)
		Exit For
	End If
Next
End Function

Private Function DecimalBinario(Texto)
If Texto="-1" Then
	DecimalBinario="B"+String(32,"1")
Else
	Do
		TextoB=""
		Modulo=0
		For z =1 To Len(Texto)
			Digito=Mid(Texto,z,1)
			Valor=CInt(Digito)+Modulo*10
			Modulo= valor Mod 2
			Cociente= valor \ 2
			TextoB=TextoB&CStr(Cociente)
		Next 
		Texto=TextoB
		DecimalBinario=CStr(Hex(Modulo))&DecimalBinario
	Loop Until  Texto=String(Len(Texto),"0")
	DecimalBinario="B"+String(32-Len(DecimalBinario),"0")+DecimalBinario
End If
End Function

  
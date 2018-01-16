DataTable.ImportSheet "C:\Users\lrua\Documents\Unified Functional Testing\Simulador\Simulador_DataDriven.xlsx",1,"Global"

SystemUtil.Run("https://www.grupobancolombia.com/wps/portal/personas/aprender-es-facil/como-manejar-dinero/endeudamiento-responsable/que-es-credito/simulador-credito-consumo")

With Browser("Simulador Crédito de Consumo").Page("Simulador Crédito de Consumo") 
	.WebList("comboTipoSimulacion").Select DataTable.Value("QueDeseaSimular","Global")
	.WebEdit("dateFechaNacimiento").Set DataTable.Value("IngresaTuFechaDeNacimiento","Global")
	.WebList("comboTipoTasa").Select DataTable.Value("ConQueTipoDeTasaQuieresTuPrestamo","Global")
	.WebList("comboTipoProducto").Select DataTable.Value("CualEsElproductoDeCreditoQueDeseasAdquirir","Global")
	.WebElement("Simula tu CréditoCalcula").Click
	.WebCheckBox("checkSeguroDesempleo").Set DataTable.Value("QuieresSeguroDeDesempleo","Global")
	.WebEdit("textPlazoInversion").Set DataTable.Value("CualEsElPlazoQueNecesitasParaTuPrestamoMeses","Global")
	.WebEdit("textValorPrestamo").Set DataTable.Value("CuantoEsElValorQueDeseasPrestar","Global")
	.WebButton("Simular").Click
	
	Dim CuotaMensual
	Set CuotaMensual = .RunScript ("document.getElementsByClassName('monto valor ng-binding')")
	 DataTable.Value("CuotaMensualMasSeguros", "Global") = CuotaMensual(5).innerText

 End With
 @@ hightlight id_;_Browser("No se puede mostrar esta").Page("Simulador Crédito de Consumo").WebButton("Simular")_;_script infofile_;_ZIP::ssf10.xml_;_
' Browser("Simulador Crédito de Consumo").Close
 
SystemUtil.Run("https://www.grupobancolombia.com/wps/portal/personas/aprender-es-facil/como-manejar-dinero/endeudamiento-responsable/credito-vs-leasing/simulador-solucion-inmobiliaria")

With Browser("Simulador Inmobiliario:").Page("Simulador Inmobiliario:")
	.WebList("combotipoFinanciacion").Select DataTable.Value("SeleccionaElTipoDeFinanciacion","Global")
	.WebElement("Selecciona el destino").Click @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebElement("Selecciona el destino")_;_script infofile_;_ZIP::ssf12.xml_;_
	.WebList("comboDestinoCredito").Select DataTable.Value("SeleccionaElDestinoDelCredito","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebList("comboDestinoCredito")_;_script infofile_;_ZIP::ssf13.xml_;_
	.WebList("comboOpcionSimular").Select DataTable.Value("SeleccionaLaOpcionASimular","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebList("comboOpcionSimular")_;_script infofile_;_ZIP::ssf14.xml_;_
	.WebList("comboPlanAmortizacion").Select DataTable.Value("SeleccionaElPlanDeAmortizacion","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebList("comboPlanAmortizacion")_;_script infofile_;_ZIP::ssf15.xml_;_
	.WebElement("Selecciona el destino_2").Click @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebElement("Selecciona el destino 2")_;_script infofile_;_ZIP::ssf16.xml_;_
	.WebEdit("textPlazoAnios").Set DataTable.Value("IngresaElPlazoEnAnos","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebEdit("textPlazoAnios")_;_script infofile_;_ZIP::ssf17.xml_;_
	.WebEdit("dateFechaNacimiento").Set DataTable.Value("IngresaTuFechaDeNacimiento","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebEdit("dateFechaNacimiento")_;_script infofile_;_ZIP::ssf18.xml_;_
	.WebList("comboDeptoColomnbia").Select DataTable.Value("SeleccionaElDepartamentoDeColombiaDondeSeEncuentraElBien","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebList("comboDeptoColomnbia")_;_script infofile_;_ZIP::ssf19.xml_;_
	.WebEdit("textValorBien").Set DataTable.Value("IngresaElValorDelBienInmueble","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebEdit("textValorBien")_;_script infofile_;_ZIP::ssf20.xml_;_
	.WebEdit("textValorPrestamo").Set DataTable.Value("IngresaElValorDelPrestamo","Global") @@ hightlight id_;_Browser("Simulador Crédito de Consumo").Page("Simulador Inmobiliario:").WebEdit("textValorPrestamo")_;_script infofile_;_ZIP::ssf21.xml_;_
	.WebButton("Simular").Click
	
	Dim CuotaHipotecario
	Set CuotaHipotecario = .RunScript ("document.getElementsByClassName('monto valor ng-binding')")
	
'	MsgBox CuotaHipotecario(19).innerText
'	MsgBox CuotaHipotecario(25).innerText
'	MsgBox CuotaHipotecario(26).innerText
	
	 DataTable.Value("CuotaMensualEnPesos", "Global") = CuotaHipotecario(19).innerText
	 DataTable.Value("SeguroDeVida", "Global") = CuotaHipotecario(25).innerText
	 DataTable.Value("SeguroDeIncendioYTerremoto", "Global") = CuotaHipotecario(26).innerText
	
End With

Para cerrar el navegador'
' Browser("Simulador Inmobiliario:").Close

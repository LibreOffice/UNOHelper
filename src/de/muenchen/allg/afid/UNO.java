/*
* Hilfsklasse zur leichteren Verwendung der UNO API.
* Dateiname: UNO.java
* Projekt  : n/a
* Funktion : Hilfsklasse zur leichteren Verwendung der UNO API.
* 
* Copyright: Landeshauptstadt M�nchen
*
* �nderungshistorie:
* Datum      | Wer | �nderungsgrund
* -------------------------------------------------------------------
* 26.04.2005 | BNK | Erstellung
* 07.07.2005 | BNK | Viele Verbesserungen
* 16.08.2005 | BNK | korrekte Dienststellenbezeichnung
* 17.08.2005 | BNK | +executeMacro()
* 17.08.2005 | BNK | +Internal-Klasse f�r interne Methoden
* 17.08.2005 | BNK | +findBrowseNodeTreeLeaf()
* 19.08.2005 | BNK | init(Object) => public, weil n�tzlich
* 19.08.2005 | BNK | +XPrintable()
* 22.08.2005 | PIT | +XNamed() 
* 22.08.2005 | PIT | +XTextContent()  
* 19.08.2005 | BNK | +XSpreadsheetDocument()
* 19.08.2005 | BNK | +XIndexAccess()
* 19.08.2005 | BNK | +XCellRange()
* 22.08.2005 | BNK | +XText()
* 22.08.2005 | PIT | +XNamed() 
* 22.08.2005 | PIT | +XTextContent()  
* 22.08.2005 | PIT | +XBookmarksSupplier()   
* 22.08.2005 | BNK | +XColumnRowRange()
* 22.08.2005 | BNK | +XTableColumns()
* 22.08.2005 | BNK | +XSpreadsheet()
* 24.08.2005 | BNK | +XCellRangeData()
* 26.08.2005 | BNK | +XBrowseNode()
* 31.08.2005 | BNK | +XScriptProvider()
*                  | +findBrowseNodeTreeLeafAndScriptProvider()
*                  | executeMacro() bekommt location Argument (diese nicht
*                  | r�ckw�rtskompatible �nderung war leider notwendig, weil die
*                  | alte Version einfach broken war und nicht innerhalb des
*                  | dokumentierten Verhaltens gefixt werden konnte.
*                  | executeGlobalMacro() durchsucht nur noch globale Makros
* 31.08.2005 | BNK | +executeMacro(macroName, args, location)
* 05.09.2005 | BNK | +XEventBroadcaster()
* 05.09.2005 | BNK | +XComponent()
* 05.09.2005 | BNK | +XTextFieldsSupplier()
* 05.09.2005 | BNK | +XModifyBroadcaster()
* 06.09.2005 | SIE | +XNameContainer()
* 06.09.2005 | SIE | +XMultiServiceFactory()
* 06.09.2005 | SIE | +XDesktop()
* 06.09.2005 | SIE | +XChangesBatch()
* 06.09.2005 | SIE | +xNameAccess()
* 06.09.2005 | BNK | TOD0 Optimierung von findBrowseNode.. hinzugef�gt
* 08.09.2005 | LUT | +xFilePicker()
* 09.09.2005 | LUT | xFilePicker() --> XFilePicker()
*                    xNameAccess() --> XNameAccess()
* ------------------------------------------------------------------- 
*
* @author D-III-ITD 5.1 Matthias S. Benkmann
* @version 1.0
* 
* */
package de.muenchen.allg.afid;

import java.util.Iterator;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.frame.FrameSearchFlag;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.script.browse.BrowseNodeFactoryViewTypes;
import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.browse.XBrowseNodeFactory;
import com.sun.star.script.provider.XScript;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.script.provider.XScriptProviderFactory;
import com.sun.star.script.provider.XScriptProviderSupplier;
import com.sun.star.sdb.XDocumentDataSource;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.sheet.XCellRangeData;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.XChangesBatch;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XURLTransformer;
import com.sun.star.view.XPrintable;

/**
 * Hilfsklasse zur leichteren Verwendung der UNO API.
 * * @author BNK
 *
 */
public class UNO {
	/**
	 * Der Haupt-ServiceManager.
	 */
	public static XMultiComponentFactory xMCF;
	/**
	 * Der Haupt-ServiceManager.
	 */
	public static XMultiServiceFactory xMSF;
	/**
	 * Das "DefaultContext" Property des Haupt-ServiceManagers. 
	 */
	public static XComponentContext defaultContext;
	/**
	 * Der globale com.sun.star.sdb.DatabaseContext.
	 */
	public static XNamingService dbContext;
	/**
	 * Ein com.sun.star.util.URLTransformer.
	 */
	public static XURLTransformer urlTransformer;
	/**
	 * Der Desktop.
	 */
	public static XDesktop desktop;
	/**
	 * Komponenten-spezifische Methoden arbeiten defaultm�ssig mit dieser
	 * Komponente. Wird von manchen Methoden ge�ndert, ist aber ansonsten der
	 * Kontroller des Programmierers �berlassen. 
	 */
	public static XComponent compo;
	
	/**
	 * Der {@link BrowseNode}, der die Wurzel des gesamten Skript-Baumes bildet.
	 */
	public static BrowseNode scriptRoot;
	
	/**
	 * Der Haupt-ScriptProvider.
	 */
	public static XScriptProvider masterScriptProvider;
	
	/**
	 * Stellt die Verbindung mit OpenOffice her unter expliziter Angabe der
	 * Verbindungsparameter. Einfacher geht's mit {@link #init()}.
	 * @param connectionString z.B. "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager"
	 * @throws Exception falls was schief geht.
	 * @see init()
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static void init(String connectionString)
	throws Exception
	{
		XComponentContext xLocalContext = Bootstrap.createInitialComponentContext(null);
		XMultiComponentFactory xLocalFactory = xLocalContext.getServiceManager();
		XUnoUrlResolver xUrlResolver = UNO.XUnoUrlResolver(xLocalFactory.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",xLocalContext));
		init(xUrlResolver.resolve(connectionString));
	}
	
	/**
	 * Stellt die Verbindung mit OpenOffice her. Die Verbindungsparameter werden
	 * automagisch ermittelt.
	 * @throws Exception falls was schief geht.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static void init()
	throws Exception
	{
		init(Bootstrap.bootstrap().getServiceManager());
	}
	
	/**
	 * Initialisiert die statischen Felder dieser Klasse ausgehend f�r eine
	 * existierende Verbindung.
	 * @param remoteServiceManager der Haupt-ServiceManager.
	 * @throws Exception falls was schief geht.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static void init(Object remoteServiceManager)
	throws Exception
	{
		xMCF = (XMultiComponentFactory)UnoRuntime.queryInterface(XMultiComponentFactory.class, remoteServiceManager);
		xMSF = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, xMCF);
		defaultContext = (XComponentContext)UnoRuntime.queryInterface(XComponentContext.class,
				((XPropertySet)UnoRuntime.queryInterface(XPropertySet.class,xMCF)).
				getPropertyValue("DefaultContext"));
		
		desktop = (XDesktop)UnoRuntime.queryInterface(XDesktop.class, xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",defaultContext));
		
		XBrowseNodeFactory masterBrowseNodeFac = (XBrowseNodeFactory) UnoRuntime.queryInterface(XBrowseNodeFactory.class ,
				defaultContext.getValueByName("/singletons/com.sun.star.script.browse.theBrowseNodeFactory"));
		scriptRoot = new BrowseNode(masterBrowseNodeFac.createView(BrowseNodeFactoryViewTypes.MACROORGANIZER));
		
		XScriptProviderFactory masterProviderFac = (XScriptProviderFactory) UnoRuntime.queryInterface(XScriptProviderFactory.class,
				defaultContext.getValueByName("/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"));
		masterScriptProvider = masterProviderFac.createScriptProvider(defaultContext);
		
		dbContext = UNO.XNamingService(UNO.createUNOService("com.sun.star.sdb.DatabaseContext"));
		urlTransformer = UNO.XURLTransformer(UNO.createUNOService("com.sun.star.util.URLTransformer"));
	}
	
	
	/**
	 * L�d ein Dokument und setzt im Erfolgsfall {@link #compo} auf das ge�ffnete Dokument.
	 * @param URL die URL des zu ladenden Dokuments, z.B. "file:///C:/temp/footest.odt"
	 *        oder "private:factory/swriter" (f�r ein leeres).
	 * @param asTemplate falls true wird das Dokument als Vorlage behandelt und ein neues
	 *        unbenanntes Dokument erzeugt.
	 * @param allowMacros  falls true wird die Ausf�hrung von Makros freigeschaltet.
	 * @return das ge�ffnete Dokument
	 * @throws com.sun.star.io.IOException
	 * @throws com.sun.star.lang.IllegalArgumentException
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static XComponent loadComponentFromURL(String URL, 
																								boolean asTemplate, boolean allowMacros)
	throws com.sun.star.io.IOException, com.sun.star.lang.IllegalArgumentException
	{
		XComponentLoader loader = UNO.XComponentLoader(desktop);
		PropertyValue[] arguments = new PropertyValue[2];
		arguments[0] = new PropertyValue();
		arguments[0].Name = "MacroExecutionMode";
		if (allowMacros)
			arguments[0].Value = new Short(MacroExecMode.ALWAYS_EXECUTE_NO_WARN);
		else
			arguments[0].Value = new Short(MacroExecMode.NEVER_EXECUTE);
		arguments[1] = new PropertyValue ();
    arguments[1].Name = "AsTemplate";
    arguments[1].Value = new Boolean(asTemplate);
    XComponent lc = loader.loadComponentFromURL(URL, "_blank", FrameSearchFlag.CREATE, arguments);
    if (lc != null) compo = lc;
		return lc; 
	}
	
	/**
	 * Ruft ein Makro auf unter expliziter Angabe der Komponente, die es zur Verf�gung
	 * stellt.
	 * 
	 * @param scriptProviderOrSupplier ist ein Objekt, das entweder {@link XScriptProvider} oder
	 *              {@link XScriptProviderSupplier} implementiert. Dies kann z.B. ein TextDocument sein.
	 *              Soll einfach nur ein Skript aus dem gesamten Skript-Baum ausgef�hrt werden,
	 *              kann die Funktion {@link #executeGlobalMacro(String, Object[])} verwendet werden, 
	 *              die diesen Parameter nicht erfordert.
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner f�r Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link IllegalArgumentException} geworfen.
	 * @param args die Argumente, die dem Makro �bergeben werden sollen.
	 * @param location eine Liste aller erlaubten locations ("application", "share", "document")
	 *        f�r das Makro. Bei der Suche wird zuerst ein case-sensitive Match in
	 *        allen gelisteten locations gesucht, bevor die case-insenstive Suche
	 *        versucht wird. Durch Verwendung der exakten Gross-/Kleinschreibung
	 *        des Makros und korrekte Ordnung der location Liste l�sst sich also
	 *        immer das richtige Makro selektieren.
	 * @throws RuntimeException wenn entweder kein passendes Makro gefunden wurde, oder
	 *              scriptProviderOrSupplier weder {@link XScriptProvider} noch
	 *              {@link XScriptProviderSupplier} implementiert.
	 * @return den R�ckgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeMacro(Object scriptProviderOrSupplier, String macroName, Object[] args, String[] location)
	{
		XScriptProvider provider = (XScriptProvider)UnoRuntime.queryInterface(XScriptProvider.class, scriptProviderOrSupplier);
		if (provider == null)
		{
			XScriptProviderSupplier supp = (XScriptProviderSupplier)UnoRuntime.queryInterface(XScriptProviderSupplier.class, scriptProviderOrSupplier);
			if (supp == null) throw new RuntimeException("�bergebenes Objekt ist weder XScriptProvider noch XScriptProviderSupplier");
			provider = supp.getScriptProvider();
		}
		
		XBrowseNode root = (XBrowseNode)UnoRuntime.queryInterface(XBrowseNode.class, provider);
		return Internal.executeMacroInternal(macroName, args, provider, root, location);
	}
	
	/**
	 * Ruft ein globales Makro auf (d,h, eines, das nicht in einem Dokument
	 * gespeichert ist). Im Falle gleichnamiger Makros hat ein Makro mit
	 * location=application Vorrang vor einem mit location=share.
	 * 
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner f�r Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link IllegalArgumentException} geworfen.
	 * @param args die Argumente, die dem Makro �bergeben werden sollen.
	 * @throws RuntimeException wenn kein passendes Makro gefunden wurde.
	 * @return den R�ckgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeGlobalMacro(String macroName, Object[] args)
	{
	  final String[] userAndShare = new String[]{"application", "share"};
	  return Internal.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), userAndShare);
	}

	/**
	 * Ruft ein Makro aus dem gesamten Makro-Baum auf.
	 * 
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner f�r Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link IllegalArgumentException} geworfen.
	 * @param args die Argumente, die dem Makro �bergeben werden sollen.
	 * @param location eine Liste aller erlaubten locations ("application", "share", "document")
	 *        f�r das Makro. Bei der Suche wird zuerst ein case-sensitive Match in
	 *        allen gelisteten locations gesucht, bevor die case-insenstive Suche
	 *        versucht wird. Durch Verwendung der exakten Gross-/Kleinschreibung
	 *        des Makros und korrekte Ordnung der location Liste l�sst sich also
	 *        immer das richtige Makro selektieren.
	 * @throws RuntimeException wenn kein passendes Makro gefunden wurde.
	 * @return den R�ckgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeMacro(String macroName, Object[] args, String[] location)
	{
	  return Internal.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), location);
	}

	
	public static class XBrowseNodeAndXScriptProvider
	{
	  public XBrowseNode XBrowseNode = null;
	  public XScriptProvider XScriptProvider = null;
	  public XBrowseNodeAndXScriptProvider(){};
	  public XBrowseNodeAndXScriptProvider(XBrowseNode xBrowseNode, XScriptProvider xScriptProvider){this.XBrowseNode = xBrowseNode; this.XScriptProvider = xScriptProvider;};
	}
	
	/**
	 * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt, dessen Name mit der
	 * Zeichefolge nameToFind endet. Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	 * @return den gefundenen Knoten oder null falls keiner gefunden.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive)
	{
	  XBrowseNodeAndXScriptProvider x = findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind, caseSensitive);
	  return x.XBrowseNode;
	}
		
	/**
	 * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt, dessen Name mit der
	 * Zeichefolge nameToFind endet. Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	 * @return den gefundenen Knoten, sowie den n�chsten Vorfahren, der XScriptProvider
	 * implementiert (oder den Knoten selbst, falls dieser XScriptProvider implementiert). 
	 * Falls kein entsprechender Knoten oder Vorfahre gefunden wurde,
	 * wird der entsprechende Wert als null geliefert.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive)
	{
	  return findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind, caseSensitive, null);	
	}

 /**
  * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt, dessen Name mit der
  * Zeichefolge nameToFind endet. Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	* @param xBrowseNode Wurzel des zu durchsuchenden Baums.
	* @param prefix wird dem Namen jedes Knoten vorangestellt. Dies wird verwendet, wenn 
	*          xBrowseNode nicht die Wurzel ist.
	* @param nameToFind  das zu suchende Namenssuffix.
	* @param caseSensitive falls true, so wird Gross-/Kleinschreibung ber�cksichtigt bei der 
	*         Suche.
  * @param Falls location nicht null ist, gelten nur Knoten als Treffer,
  *        die ein "URI" Property haben, das eine location enth�lt die
  *        einem String in der <code>location</code> Liste entspricht.
	*        Typischerweise werden f�r <code>location</code> "application" 
	*        oder "share" �bergeben.
	* @return den gefundenen Knoten, sowie den n�chsten Vorfahren, der XScriptProvider
	* implementiert. Falls kein entsprechender Knoten oder Vorfahre gefunden wurde,
	* wird der entsprechende Wert als null geliefert.
	* @author Matthias Benkmann (D-III-ITD 5.1)
	*/
	public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive, String[] location)
	{ //T
		final String[] noLoc = new String[]{null};
		if (location == null) location = noLoc;
		for (int i = 0; i < location.length; ++i)
		{
	    XBrowseNodeAndXScriptProvider o = Internal.findBrowseNodeTreeLeafAndScriptProvider(new BrowseNode(xBrowseNode), prefix, nameToFind, caseSensitive, location[i], null);
	    if (o.XBrowseNode != null) return o;
		}
		
		return new XBrowseNodeAndXScriptProvider(null,null);
	}
	
	/**
	 * Liefert den Wert von Property propName des Objekts o zur�ck.
	 * @return den Wert des Propertys oder <code>null</code>, falls o
	 * entweder nicht das XPropertySet Interface implementiert, oder kein
	 * Property names propName hat.
	 */
	public static Object getProperty(Object o, String propName)
	{
		Object ret = null;
		try {
			XPropertySet props = UNO.XPropertySet(o);
			if (props == null) return null;
			ret = props.getPropertyValue(propName);
		} catch (UnknownPropertyException e) {
		  return null;
		} catch (WrappedTargetException e) {
				e.printStackTrace();
				System.exit(0);
		}
		return ret;
	}
	
	/**
	 * Setzt das Property propName des Objekts o auf den Wert propVal und liefert
	 * den neuen Wert zur�ck. Falls o kein XPropertySet implementiert, oder
	 * das Property propName nicht gelesen werden kann (z.B. weil o diese
	 * Property nicht besitzt), so wird null zur�ckgeliefert.
	 * Zu beachten ist, dass es m�glich ist, dass der zur�ckgelieferte Wert
	 * nicht propVal und auch nicht null ist. Dies geschieht insbesondere,
	 * wenn ein Event Handler sein Veto gegen die �nderung eingelegt hat.
	 * @param o das Objekt, dessen Property zu �ndern ist.
	 * @param propName der Name des zu �ndernden Properties.
	 * @param propVal der neue Wert.
	 * @return der Wert des Propertys nach der (versuchten) �nderung oder null,
	 *         falls der Wert des Propertys nicht mal lesbar ist.
	 * @author bnk
	 */
	public static Object setProperty(Object o, String propName, Object propVal)
	{
		Object ret = null;
		try {
			XPropertySet props = UNO.XPropertySet(o);
			if (props == null) return null;
			try{
				props.setPropertyValue(propName, propVal);
			} catch(Exception x){}
			ret = props.getPropertyValue(propName);
		} catch (UnknownPropertyException e) {
		  return null;
		} catch (WrappedTargetException e) {
				e.printStackTrace();
				System.exit(0);
		}
		return ret;
	}

	
	/**
	 * Erzeugt einen Dienst im Haupt-Servicemanager mit dem DefaultContext.
	 * @param serviceName Name des zu erzeugenden Dienstes
	 * @return ein Objekt, das den Dienst anbietet, oder null falls Fehler.
	 * @author bnk
	 */
	public static Object createUNOService(String serviceName)
	{
		try{
		  return xMCF.createInstanceWithContext(serviceName, defaultContext);
		}catch(Exception x)
		{
			return null;
		}
	}

	/** Holt {@link XSingleServiceFactory} Interface von o.*/
	public static XSingleServiceFactory XSingleServiceFactory(Object o)
	{
		return (XSingleServiceFactory)UnoRuntime.queryInterface(XSingleServiceFactory.class,o);
	}

	/** Holt {@link XStorable} Interface von o.*/
	public static XStorable XStorable(Object o)
	{
		return (XStorable)UnoRuntime.queryInterface(XStorable.class,o);
	}
	
	/** Holt {@link XTextFieldsSupplier} Interface von o.*/
	public static XTextFieldsSupplier XTextFieldsSupplier(Object o)
	{
		return (XTextFieldsSupplier)UnoRuntime.queryInterface(XTextFieldsSupplier.class,o);
	}
	
	/** Holt {@link XComponent} Interface von o.*/
	public static XComponent XComponent(Object o)
	{
		return (XComponent)UnoRuntime.queryInterface(XComponent.class,o);
	}
	
	/** Holt {@link XEventBroadcaster} Interface von o.*/
	public static XEventBroadcaster XEventBroadcaster(Object o)
	{
		return (XEventBroadcaster)UnoRuntime.queryInterface(XEventBroadcaster.class,o);
	}
	
	/** Holt {@link XBrowseNode} Interface von o.*/
	public static XBrowseNode XBrowseNode(Object o)
	{
		return (XBrowseNode)UnoRuntime.queryInterface(XBrowseNode.class,o);
	}
	
	/** Holt {@link XCellRangeData} Interface von o.*/
	public static XCellRangeData XCellRangeData(Object o)
	{
		return (XCellRangeData)UnoRuntime.queryInterface(XCellRangeData.class,o);
	}
	
 	/** Holt {@link XTableColumns} Interface von o.*/
	public static XTableColumns XTableColumns(Object o)
	{
		return (XTableColumns)UnoRuntime.queryInterface(XTableColumns.class,o);
	}

	
	/** Holt {@link XColumnRowRange} Interface von o.*/
	public static XColumnRowRange XColumnRowRange(Object o)
	{
		return (XColumnRowRange)UnoRuntime.queryInterface(XColumnRowRange.class,o);
	}
	
	/** Holt {@link XIndexAccess} Interface von o.*/
	public static XIndexAccess XIndexAccess(Object o)
	{
		return (XIndexAccess)UnoRuntime.queryInterface(XIndexAccess.class,o);
	}
	
	/** Holt {@link XCellRange} Interface von o.*/
	public static XCellRange XCellRange(Object o)
	{
		return (XCellRange)UnoRuntime.queryInterface(XCellRange.class,o);
	}

	/** Holt {@link XModifiable} Interface von o.*/
	public static XModifiable XModifiable(Object o)
	{
		return (XModifiable)UnoRuntime.queryInterface(XModifiable.class,o);
	}

	
	/** Holt {@link XDocumentDataSource} Interface von o.*/
	public static XDocumentDataSource XDocumentDataSource(Object o)
	{
		return (XDocumentDataSource)UnoRuntime.queryInterface(XDocumentDataSource.class,o);
	}

	
	/** Holt {@link XNamingService} Interface von o.*/
	public static XNamingService XNamingService(Object o)
	{
		return (XNamingService)UnoRuntime.queryInterface(XNamingService.class,o);
	}
	
	/** Holt {@link XNamed} Interface von o.*/
	public static XNamed XNamed(Object o)
	{
		return (XNamed)UnoRuntime.queryInterface(XNamed.class,o);
	}	

	
	/** Holt {@link XUnoUrlResolver} Interface von o.*/
	public static XUnoUrlResolver XUnoUrlResolver(Object o)
	{
		return (XUnoUrlResolver)UnoRuntime.queryInterface(XUnoUrlResolver.class,o);
	}
	
	/** Holt {@link XPropertySet} Interface von o.*/
	public static XPropertySet XPropertySet(Object o)
	{
		return (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class,o);
	}

	/** Holt {@link XModel} Interface von o.*/
	public static XModel XModel(Object o)
	{
		return (XModel)UnoRuntime.queryInterface(XModel.class,o);
	}
	
	/** Holt {@link XModifyBroadcaster} Interface von o.*/
	public static XModifyBroadcaster XModifyBroadcaster(Object o)
	{
		return (XModifyBroadcaster)UnoRuntime.queryInterface(XModifyBroadcaster.class,o);
	}
	
	/** Holt {@link XScriptProvider} Interface von o.*/
	public static XScriptProvider XScriptProvider(Object o)
	{
		return (XScriptProvider)UnoRuntime.queryInterface(XScriptProvider.class,o);
	}
	
	/** Holt {@link XSpreadsheet} Interface von o.*/
	public static XSpreadsheet XSpreadsheet(Object o)
	{
		return (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class,o);
	}
	
	/** Holt {@link XText} Interface von o.*/
	public static XText XText(Object o)
	{
		return (XText)UnoRuntime.queryInterface(XText.class,o);
	}
	
	/** Holt {@link XSpreadsheetDocument} Interface von o.*/
	public static XSpreadsheetDocument XSpreadsheetDocument(Object o)
	{
		return (XSpreadsheetDocument)UnoRuntime.queryInterface(XSpreadsheetDocument.class,o);
	}

	/** Holt {@link XPrintable} Interface von o.*/
	public static XPrintable XPrintable(Object o)
	{
		return (XPrintable)UnoRuntime.queryInterface(XPrintable.class,o);
	}

	/** Holt {@link XTextDocument} Interface von o.*/
	public static XTextDocument XTextDocument(Object o)
	{
		return (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class,o);
	}
	
	/** Holt {@link XTextContent} Interface von o.*/
	public static XTextContent XTextContent(Object o)
	{
		return (XTextContent)UnoRuntime.queryInterface(XTextContent.class,o);
	}	

	
	/** Holt {@link XDispatchProvider} Interface von o.*/
	public static XDispatchProvider XDispatchProvider(Object o)
	{
		return (XDispatchProvider)UnoRuntime.queryInterface(XDispatchProvider.class,o);
	}

	
	/** Holt {@link XURLTransformer} Interface von o.*/
	public static XURLTransformer XURLTransformer(Object o)
	{
		return (XURLTransformer)UnoRuntime.queryInterface(XURLTransformer.class,o);
	}

	
	/** Holt {@link XComponentLoader} Interface von o.*/
	public static XComponentLoader XComponentLoader(Object o)
	{
		return (XComponentLoader)UnoRuntime.queryInterface(XComponentLoader.class,o);
	}

	/** Holt {@link XBrowseNodeFactory} Interface von o.*/
	public static XBrowseNodeFactory XBrowseNodeFactory(Object o)
	{
		return (XBrowseNodeFactory)UnoRuntime.queryInterface(XBrowseNodeFactory.class,o);
	}
	
	/** Holt {@link XBookmarksSupplier} Interface von o.*/
	public static XBookmarksSupplier XBookmarksSupplier(Object o)
	{
		return (XBookmarksSupplier)UnoRuntime.queryInterface(XBookmarksSupplier.class,o);
	}	

	/** Holt {@link XNameContainer} Interface von o.*/
	public static XNameContainer XNameContainer(Object o)
	{
		return (XNameContainer)UnoRuntime.queryInterface(XNameContainer.class,o);
	}	
	
	/** Holt {@link XMultiServiceFactory} Interface von o.*/
	public static XMultiServiceFactory XMultiServiceFactory(Object o)
	{
		return (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class,o);
	}	
	
	/** Holt {@link XDesktop} Interface von o.*/
	public static XDesktop XDesktop(Object o)
	{
		return (XDesktop)UnoRuntime.queryInterface(XDesktop.class,o);
	}	
	
	/** Holt {@link XChangesBatch} Interface von o.*/
	public static XChangesBatch XChangesBatch(Object o)
	{
		return (XChangesBatch)UnoRuntime.queryInterface(XChangesBatch.class,o);
	}	
	
	/** Holt {@link XNameAccess} Interface von o.*/
	public static XNameAccess XNameAccess(Object o)
	{
		return (XNameAccess)UnoRuntime.queryInterface(XNameAccess.class,o);
	}	

	/** Holt {@link XFilePicker} Interface von o.*/
	public static XFilePicker XFilePicker(Object o)
	{
		return (XFilePicker)UnoRuntime.queryInterface(XFilePicker.class,o);
	}	
	
	// ACHTUNG: Interface-Methoden fangen hier mit einem grossen X an!
  /**
   * Interne Funktionen
   */
	private static class Internal
	{

	  /**
	   * siehe {@link UNO#findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String, XScriptProvider)}
	   * @param xScriptProvider der zuletzt gesehene xScriptProvider 
	   * @author bnk
	   */
	  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(BrowseNode node, String prefix, String nameToFind, boolean caseSensitive, String location, XScriptProvider xScriptProvider)
		{ //T
	    /* TODO M�gliche Optimierung, um Geschwindigkeit zu steigern:
	     * Statt Tiefensuche eine Breitensuche unter Bevorzugung von �sten mit
	     * der Eigenschaft, dass ein Suffix des Astbezeichners, das mit einem 
	     * Punkt beginnt ein Prefix ist des zu suchenden Bezeichners.
	     * Dadurch sollten Anfragen, die einen m�glichst pr�zisen Pfad liefern
	     * (inklusive Bibliothek und Modul) wesentlich schneller werden.
	     */
	    String name = node.getName();
			if (!prefix.equals(""))	name = prefix + "." + name; 
			
			if (!caseSensitive) 
			{ 
				name = name.toLowerCase(); 
				nameToFind = nameToFind.toLowerCase(); 
			}
			
			XScriptProvider xsc = (XScriptProvider)node.as(XScriptProvider.class);
			if (xsc != null) xScriptProvider = xsc;
			
			Iterator iter = node.children();
			
			if (name.endsWith(nameToFind) && !iter.hasNext() && checkLocation(node.unwrap(), location)) 
			  return new XBrowseNodeAndXScriptProvider(node.unwrap(), xScriptProvider);
						
			while (iter.hasNext())
			{
				BrowseNode child = (BrowseNode)iter.next();
				
				XBrowseNodeAndXScriptProvider o = findBrowseNodeTreeLeafAndScriptProvider(child, name, nameToFind, caseSensitive, location, xScriptProvider);
        if (o.XBrowseNode != null) return o;
			}
			
			return new XBrowseNodeAndXScriptProvider(null, null);
		}

	  /**
	   * @author bnk
	   */
	  private static boolean checkLocation(XBrowseNode node, String location)
	  {
	    if (location == null) return true;
	    String uri = (String)UNO.getProperty(node, "URI");
	    if (uri == null) return false;
	    if (uri.indexOf("?location="+location) >= 0 ||
	        uri.indexOf("&location="+location) >= 0) return true;
	    return false;
	  }
	  
    /** 
     * Wenn <code>provider = null</code>, so wird versucht, 
     * einen passenden Provider zu finden.  
     * @author bnk
     */
    private static Object executeMacroInternal(String macroName, Object[] args, XScriptProvider provider, XBrowseNode root, String[] location)
    {
      XBrowseNodeAndXScriptProvider o = null;
      
      o = UNO.findBrowseNodeTreeLeafAndScriptProvider(root, "", "." + macroName, true, location);
      if (o.XBrowseNode == null)
        o = UNO.findBrowseNodeTreeLeafAndScriptProvider(root, "", "." + macroName, false, location);
      if (o.XBrowseNode == null)
        o = UNO.findBrowseNodeTreeLeafAndScriptProvider(root, "", macroName, true, location);
      if (o.XBrowseNode == null)
        o = UNO.findBrowseNodeTreeLeafAndScriptProvider(root, "", macroName, false, location);
    	
      if (provider == null) provider = o.XScriptProvider;
    	
      XScript script;
    	
    	try{
    	  String uri = (String)UNO.getProperty(o.XBrowseNode, "URI");
    		script = provider.getScript(uri);
    	}
    	catch(Exception x)
    	{
    		throw new RuntimeException("Objekt "+macroName+" nicht gefunden oder ist kein Skript");
    	}
    	
    	short[][] aOutParamIndex = new short[][]{new short[0]};
    	Object[][] aOutParam = new Object[][]{new Object[0]};
    	try{
    		Object retval = script.invoke(args, aOutParamIndex, aOutParam);
    		return retval;
    	} catch(Exception x)
    	{
    		x.printStackTrace();
    		throw new RuntimeException("Fehler bei invoke() von Makro "+macroName);
    	}
    }
	}
}

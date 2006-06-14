/*
* Hilfsklasse zur leichteren Verwendung der UNO API.
* Dateiname: UNO.java
* Projekt  : n/a
* Funktion : Hilfsklasse zur leichteren Verwendung der UNO API.
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Datum      | Wer | Änderungsgrund
* -------------------------------------------------------------------
* 26.04.2005 | BNK | Erstellung
* 07.07.2005 | BNK | Viele Verbesserungen
* 16.08.2005 | BNK | korrekte Dienststellenbezeichnung
* 17.08.2005 | BNK | +executeMacro()
* 17.08.2005 | BNK | +Internal-Klasse für interne Methoden
* 17.08.2005 | BNK | +findBrowseNodeTreeLeaf()
* 19.08.2005 | BNK | init(Object) => public, weil nützlich
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
*                  | rückwärtskompatible Änderung war leider notwendig, weil die
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
* 06.09.2005 | BNK | TOD0 Optimierung von findBrowseNode.. hinzugefügt
* 08.09.2005 | LUT | +xFilePicker()
* 09.09.2005 | LUT | xFilePicker() --> XFilePicker()
*            |     | xNameAccess() --> XNameAccess()
* 13.09.2005 | LUT | +XToolkit()
* 13.09.2005 | LUT | +XWindow()
*                    +XWindowPeer()
*                    +XToolbarController()
* 13.09.2005 | BNK | findBrowseNode...() Doku geändert, so dass sie sagt, dass
*            |     | nur Blätter vom Typ SCRIPT gesucht werden.
* 13.09.2005 | BNK | +Internal.findBrowseNodeTreeLeavesAndScriptProviders()
*            |     | UNO.findBrowseNode... geändert zur Verwdg. der obigen Fkt.
* 14.09.2005 | BNK | Bugs gefixt
* 02.01.2005 | BNK | +XEnumerationAccess()
* 31.01.2006 | BNK | +XViewSettingsSupplier()
* 31.01.2006 | BNK | +XWindow2()
* 01.02.2006 | BNK | +XMultiPropertySet()
*                  | +XLayoutManager()
*                  | +XCloseable()
*                  | +XTopWindow()
* 18.05.2006 | BNK | +XDocumentInsertable()
*                  | +XTextField()
* 09.06.2006 | LUT | getPropertyValue und setPropertyValue dürfen nicht System.exit(0) aufrufen,
*                    wenn eine WrappedTargetException auftrat !
* 14.06.2006 | LUT | +XServiceInfo()
*                  | +supportsService(...)
* ------------------------------------------------------------------- 
*
* @author D-III-ITD 5.1 Matthias S. Benkmann
* @version 1.0
* 
* */
package de.muenchen.allg.afid;

import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.ListIterator;
import java.util.Vector;

import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindow2;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.frame.FrameSearchFlag;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.frame.XToolbarController;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.script.browse.BrowseNodeFactoryViewTypes;
import com.sun.star.script.browse.BrowseNodeTypes;
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
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.XChangesBatch;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XURLTransformer;
import com.sun.star.view.XPrintable;
import com.sun.star.view.XViewSettingsSupplier;

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
	 * Komponenten-spezifische Methoden arbeiten defaultmässig mit dieser
	 * Komponente. Wird von manchen Methoden geändert, ist aber ansonsten der
	 * Kontroller des Programmierers überlassen. 
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
	 * Initialisiert die statischen Felder dieser Klasse ausgehend für eine
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
	 * Läd ein Dokument und setzt im Erfolgsfall {@link #compo} auf das geöffnete Dokument.
	 * @param URL die URL des zu ladenden Dokuments, z.B. "file:///C:/temp/footest.odt"
	 *        oder "private:factory/swriter" (für ein leeres).
	 * @param asTemplate falls true wird das Dokument als Vorlage behandelt und ein neues
	 *        unbenanntes Dokument erzeugt.
	 * @param allowMacros  falls true wird die Ausführung von Makros freigeschaltet.
	 * @return das geöffnete Dokument
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
	 * Ruft ein Makro auf unter expliziter Angabe der Komponente, die es zur Verfügung
	 * stellt.
	 * 
	 * @param scriptProviderOrSupplier ist ein Objekt, das entweder {@link XScriptProvider} oder
	 *              {@link XScriptProviderSupplier} implementiert. Dies kann z.B. ein TextDocument sein.
	 *              Soll einfach nur ein Skript aus dem gesamten Skript-Baum ausgeführt werden,
	 *              kann die Funktion {@link #executeGlobalMacro(String, Object[])} verwendet werden, 
	 *              die diesen Parameter nicht erfordert.
	 *              ACHTUNG! Es wird nicht zwangsweise der übergebene scriptProviderOrSupplier
	 *              verwendet um das Skript auszuführen. Er stellt nur den Einstieg in
	 *              den Skript-Baum dar.
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner für Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link RuntimeException} geworfen.
	 * @param args die Argumente, die dem Makro übergeben werden sollen.
	 * @param location eine Liste aller erlaubten locations ("application", "share", "document")
	 *        für das Makro. Bei der Suche wird zuerst ein case-sensitive Match in
	 *        allen gelisteten locations gesucht, bevor die case-insensitive Suche
	 *        versucht wird. Durch Verwendung der exakten Gross-/Kleinschreibung
	 *        des Makros und korrekte Ordnung der location Liste lässt sich also
	 *        immer das richtige Makro selektieren.
	 * @throws RuntimeException wenn entweder kein passendes Makro gefunden wurde, oder
	 *              scriptProviderOrSupplier weder {@link XScriptProvider} noch
	 *              {@link XScriptProviderSupplier} implementiert.
	 * @return den Rückgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeMacro(Object scriptProviderOrSupplier, String macroName, Object[] args, String[] location)
	{
		XScriptProvider provider = (XScriptProvider)UnoRuntime.queryInterface(XScriptProvider.class, scriptProviderOrSupplier);
		if (provider == null)
		{
			XScriptProviderSupplier supp = (XScriptProviderSupplier)UnoRuntime.queryInterface(XScriptProviderSupplier.class, scriptProviderOrSupplier);
			if (supp == null) throw new RuntimeException("Übergebenes Objekt ist weder XScriptProvider noch XScriptProviderSupplier");
			provider = supp.getScriptProvider();
		}
		
		XBrowseNode root = (XBrowseNode)UnoRuntime.queryInterface(XBrowseNode.class, provider);
		/*
		 * Wir übergeben NICHT provider als drittes Argument, sondern lassen
		 * Internal.executeMacroInternal den provider selbst bestimmen.
		 * Das hat keinen besonderen Grund. Es erscheint einfach nur etwas
		 * robuster, den "nächstgelegenen" ScriptProvider zu verwenden.
		 */
		return Internal.executeMacroInternal(macroName, args, null, root, location);
	}
	
	/**
	 * Ruft ein globales Makro auf (d,h, eines, das nicht in einem Dokument
	 * gespeichert ist). Im Falle gleichnamiger Makros hat ein Makro mit
	 * location=application Vorrang vor einem mit location=share.
	 * 
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner für Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link RuntimeException} geworfen.
	 * @param args die Argumente, die dem Makro übergeben werden sollen.
	 * @throws RuntimeException wenn kein passendes Makro gefunden wurde.
	 * @return den Rückgabewert des Makros.
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
	 *              Bezeichner für Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link RuntimeException} geworfen.
	 * @param args die Argumente, die dem Makro übergeben werden sollen.
	 * @param location eine Liste aller erlaubten locations ("application", "share", "document")
	 *        für das Makro. Bei der Suche wird zuerst ein case-sensitive Match in
	 *        allen gelisteten locations gesucht, bevor die case-insenstive Suche
	 *        versucht wird. Durch Verwendung der exakten Gross-/Kleinschreibung
	 *        des Makros und korrekte Ordnung der location Liste lässt sich also
	 *        immer das richtige Makro selektieren.
	 * @throws RuntimeException wenn kein passendes Makro gefunden wurde.
	 * @return den Rückgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeMacro(String macroName, Object[] args, String[] location)
	{
	  return Internal.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), location);
	}

	/**
     * Diese Methode prüft, ob es sich bei dem übergebenen Objekt service um ein
     * UNO-Service mit dem Namen serviceName handelt und liefert true zurück,
     * wenn das Objekt das XServiceInfo-Interface und den gesuchten Service
     * implementiert, ansonsten wird false zurückgegeben.
     * 
     * @param service
     *            Das zu prüfende Service-Objekt
     * @param serviceName
     *            der voll-qualifizierte Service-Name des services.
     * @return true, wenn das Objekt das XServiceInfo-Interface und den
     *         gesuchten Service implementiert, ansonsten false.
     * @author Christoph Lutz (D-III-ITD 5.1)
     */
    public static boolean supportsService(Object service, String serviceName) {
        if (UNO.XServiceInfo(service) != null)
            return UNO.XServiceInfo(service).supportsService(serviceName);
        return false;
    }
	
	public static class XBrowseNodeAndXScriptProvider
	{
	  public XBrowseNode XBrowseNode = null;
	  public XScriptProvider XScriptProvider = null;
	  public XBrowseNodeAndXScriptProvider(){};
	  public XBrowseNodeAndXScriptProvider(XBrowseNode xBrowseNode, XScriptProvider xScriptProvider){this.XBrowseNode = xBrowseNode; this.XScriptProvider = xScriptProvider;};
	}
	
	/**
	 * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT, dessen Name 
	 * nameToFind ist (kann durch "." abgetrennte Pfadangabe im Skript-Baum enthalten).
	 * Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	 * @return den gefundenen Knoten oder null falls keiner gefunden.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive)
	{
	  XBrowseNodeAndXScriptProvider x = findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind, caseSensitive);
	  return x.XBrowseNode;
	}
	
	/**
	 * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT, dessen Name
	 * nameToFind ist (kann durch "." abgetrennte Pfadangabe im Skript-Baum enthalten). 
	 * Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	 * @return den gefundenen Knoten, sowie den nächsten Vorfahren, der XScriptProvider
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
  * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT, dessen Name
  * nameToFind ist (kann durch "." abgetrennte Pfadangabe im Skript-Baum enthalten).
  *  Siehe {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}. 
	* @param xBrowseNode Wurzel des zu durchsuchenden Baums.
	* @param prefix wird dem Namen jedes Knoten vorangestellt. Dies wird verwendet, wenn 
	*          xBrowseNode nicht die Wurzel ist.
	* @param nameToFind  der zu suchende Name.
	* @param caseSensitive falls true, so wird Gross-/Kleinschreibung berücksichtigt bei der 
	*         Suche.
  * @param location Es gelten nur Knoten als Treffer,
  *        die ein "URI" Property haben, das eine location enthält die
  *        einem String in der <code>location</code> Liste entspricht.
	*        Mögliche locations sind "document", "application" und "share".
	*        Falls <code>location==null</code>, so wird {"document", "application", "share"}
	*        angenommen.
	* @return den gefundenen Knoten, sowie den nächsten Vorfahren, der XScriptProvider
	* implementiert. Falls kein entsprechender Knoten oder Vorfahre gefunden wurde,
	* wird der entsprechende Wert als null geliefert.
	* @author Matthias Benkmann (D-III-ITD 5.1)
	*/
	public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive, String[] location)
	{ //T
		final String[] noLoc = new String[]{"document", "application", "share"};
		if (location == null) location = noLoc;
		List found = new LinkedList();
		List prefixVec = new Vector();
		List prefixLCVec = new Vector();
		String[] prefixArr = prefix.split("\\.");
		for (int i = 0; i < prefixArr.length; ++i)
		  if (!prefixArr[i].equals("")) 
		  {
		    prefixVec.add(prefixArr[i]);
		    prefixLCVec.add(prefixArr[i].toLowerCase());
		  }
		
		String[] nameToFindArrPre = nameToFind.split("\\.");
		int i1 = 0;
		while (i1 < nameToFindArrPre.length && nameToFindArrPre[i1].equals("")) ++i1;
		int i2 = nameToFindArrPre.length - 1;
		while (i2 >= i1 && nameToFindArrPre[i2].equals("")) --i2;
		++i2;
		String[] nameToFindArr = new String[i2 - i1];
		String[] nameToFindLCArr = new String[i2 - i1];
	  for (int i = 0; i < nameToFindArr.length; ++i)
	  {
	    nameToFindArr[i] = nameToFindArrPre[i1];
	    nameToFindLCArr[i] = nameToFindArrPre[i1].toLowerCase();
	    ++i1;
	  }
	  
		Internal.findBrowseNodeTreeLeavesAndScriptProviders(new BrowseNode(xBrowseNode), prefixVec, prefixLCVec, nameToFindArr, nameToFindLCArr, location, null, found);
		
		if (found.isEmpty())
		  return new XBrowseNodeAndXScriptProvider(null,null);
		
		Internal.FindNode findNode = (Internal.FindNode)found.get(0);
		
		if (caseSensitive && !findNode.isCaseCorrect)
		  return new XBrowseNodeAndXScriptProvider(null,null);
		
		return new XBrowseNodeAndXScriptProvider(findNode.XBrowseNode, findNode.XScriptProvider);
	}
	
	/**
	 * Liefert den Wert von Property propName des Objekts o zurück.
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
		} catch (WrappedTargetException e) {
		}
		return ret;
	}
	
	/**
	 * Setzt das Property propName des Objekts o auf den Wert propVal und liefert
	 * den neuen Wert zurück. Falls o kein XPropertySet implementiert, oder
	 * das Property propName nicht gelesen werden kann (z.B. weil o diese
	 * Property nicht besitzt), so wird null zurückgeliefert.
	 * Zu beachten ist, dass es möglich ist, dass der zurückgelieferte Wert
	 * nicht propVal und auch nicht null ist. Dies geschieht insbesondere,
	 * wenn ein Event Handler sein Veto gegen die Änderung eingelegt hat.
	 * @param o das Objekt, dessen Property zu ändern ist.
	 * @param propName der Name des zu ändernden Properties.
	 * @param propVal der neue Wert.
	 * @return der Wert des Propertys nach der (versuchten) Änderung oder null,
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
		} catch (WrappedTargetException e) {
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

    /** Holt {@link XViewSettingsSupplier} Interface von o.*/
    public static XViewSettingsSupplier XViewSettingsSupplier(Object o)
    {
        return (XViewSettingsSupplier)UnoRuntime.queryInterface(XViewSettingsSupplier.class,o);
    }
    
    
	/** Holt {@link XStorable} Interface von o.*/
	public static XStorable XStorable(Object o)
	{
		return (XStorable)UnoRuntime.queryInterface(XStorable.class,o);
	}
    

    /** Holt {@link XTopWindow} Interface von o.*/
    public static XTopWindow XTopWindow(Object o)
    {
        return (XTopWindow)UnoRuntime.queryInterface(XTopWindow.class,o);
    }
    
    /** Holt {@link XCloseable} Interface von o.*/
    public static XCloseable XCloseable(Object o)
    {
        return (XCloseable)UnoRuntime.queryInterface(XCloseable.class,o);
    }
    
    /** Holt {@link XWindow2} Interface von o.*/
    public static XWindow2 XWindow2(Object o)
    {
        return (XWindow2)UnoRuntime.queryInterface(XWindow2.class,o);
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
	
    /** Holt {@link XMultiPropertySet} Interface von o.*/
    public static XMultiPropertySet XMultiPropertySet(Object o)
    {
        return (XMultiPropertySet)UnoRuntime.queryInterface(XMultiPropertySet.class,o);
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
    
    /** Holt {@link XTextField} Interface von o.*/
    public static XTextField XTextField(Object o)
    {
        return (XTextField)UnoRuntime.queryInterface(XTextField.class,o);
    }
    
    /** Holt {@link XDocumentInsertable} Interface von o.*/
    public static XDocumentInsertable XDocumentInsertable(Object o)
    {
        return (XDocumentInsertable)UnoRuntime.queryInterface(XDocumentInsertable.class,o);
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
    
    /** Holt {@link XLayoutManager} Interface von o.*/
    public static XLayoutManager XLayoutManager(Object o)
    {
        return (XLayoutManager)UnoRuntime.queryInterface(XLayoutManager.class,o);
    }
    
    /** Holt {@link XEnumerationAccess} Interface von o.*/
    public static XEnumerationAccess XEnumerationAccess(Object o)
    {
        return (XEnumerationAccess)UnoRuntime.queryInterface(XEnumerationAccess.class,o);
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

	/** Holt {@link XToolkit} Interface von o.*/
	public static XToolkit XToolkit(Object o)
	{
		return (XToolkit)UnoRuntime.queryInterface(XToolkit.class,o);
	}	
	
	/** Holt {@link XWindow} Interface von o.*/
	public static XWindow XWindow(Object o)
	{
		return (XWindow)UnoRuntime.queryInterface(XWindow.class,o);
	}	
	
	/** Holt {@link XToolkit} Interface von o.*/
	public static XWindowPeer XWindowPeer(Object o)
	{
		return (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class,o);
	}	
	
    /** Holt {@link XToolbarController} Interface von o.*/
    public static XToolbarController XToolbarController(Object o)
    {
        return (XToolbarController)UnoRuntime.queryInterface(XToolbarController.class,o);
    }   

    /** Holt {@link com.sun.star.text.XParagraphCursor} Interface von o.*/
    public static XParagraphCursor XParagraphCursor(Object o)
    {
        return (XParagraphCursor)UnoRuntime.queryInterface(XParagraphCursor.class,o);
    }   

    /** Holt {@link com.sun.star.lang.XServiceInf} Interface von o.*/
    public static XServiceInfo XServiceInfo(Object o)
    {
        return (XServiceInfo)UnoRuntime.queryInterface(XServiceInfo.class,o);
    }   

	// ACHTUNG: Interface-Methoden fangen hier mit einem grossen X an!
  /**
   * Interne Funktionen
   */
	private static class Internal
	{

	  public static class FindNode extends UNO.XBrowseNodeAndXScriptProvider
	  {
	    public String location;
	    public boolean isCaseCorrect;
	    public FindNode(XBrowseNode xBrowseNode, XScriptProvider xScriptProvider, String location, boolean isCaseCorrect) 
	    {
	      super(xBrowseNode, xScriptProvider);
	      this.location = location;
	      this.isCaseCorrect = isCaseCorrect;
	    }
      
	    /**
	     * Returns true if <code>this</code> isCaseCorrect and fn2 is not,
	     * or both have the same isCaseCorrect but 
	     * <code>this.location</code> occurs earlier in locations than
	     * <code>fn2.location</code>.
       * @author bnk
       */
      public boolean betterMatchThan(Object fn2, String[] locations)
      {
        FindNode f2 = (FindNode)fn2;
        if (this.isCaseCorrect && !f2.isCaseCorrect) return true;
        if (f2.isCaseCorrect && !this.isCaseCorrect) return false;
        int i = 0;
        int i2 = 0;
        while (i < locations.length && !locations[i].equals(this.location)) ++i; 
        while (i2 < locations.length && !locations[i2].equals(f2.location)) ++i2;
        return i < i2;
      }
	  }
	  
	  /**
	   * siehe {@link UNO#findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String[]))}
	   * @param xScriptProvider der zuletzt gesehene xScriptProvider
	   * @param nameToFind der zu suchende Name in seine Bestandteile zwischen den Punkten
	   *        zerlegt.
	   * @param nameToFindLC wie nameToFind aber alles lowercase.
	   * @param prefix das Prefix in seine Bestandteile zwischen den Punkten
	   *        zerlegt.
	   * @param prefixLC wie prefix aber alles lowercase.
	   * @param found Liste von {@link FindNode}s mit dem Ergebnis der Suche 
	   *        (anfangs leere Liste übergeben).
	   *        Die Sortierung ist so, dass zuerst alle case-sensitive Matches 
	   *        (also exakte Matches) aufgeführt sind, sortiert gemäss location
	   *        und dann alle case-insensitive Matches sortiert gemäss location.
	   *        Falls <code>location == null</code>, so wird nur nach case-sensitive
	   *        und case-insenstive sortiert, innerhalb dieser Gruppen jedoch
	   *        nicht mehr.
	   * @return die Anzahl der Rekursionsstufen, die beendet werden sollen.
	   *         Zum Beispiel heisst ein Rückgabewert von 1, dass die aufrufende
	   *         Funktion ein <code>return 0</code> machen soll.
	   *         
	   * @author bnk
	   */
	  public static int findBrowseNodeTreeLeavesAndScriptProviders(BrowseNode node, List prefix, List prefixLC, String[] nameToFind, String[] nameToFindLC, String[] location, XScriptProvider xScriptProvider, List found)
		{ //T
	    String name = node.getName();
	    String nameLC = name.toLowerCase();
	    
	    XScriptProvider xsc = (XScriptProvider)node.as(XScriptProvider.class);
			if (xsc != null) xScriptProvider = xsc;
	    	    
			Iterator iter = node.children();
			if (!iter.hasNext())
			{
			  /*
			   * Falls der Knoten nicht vom Typ SCRIPT ist, interessiert er uns
			   * nicht. Auch wenn wir davon ausgehen können, dass alle Geschwister
			   * ebenfalls keine SCRIPTS sind, dürfen wir nicht mehrere Stufen nach
			   * oben gehen, da die Geschwister CONTAINER sein können.
			   */
			  if (node.getType() != BrowseNodeTypes.SCRIPT) return 0;
			  
			  /* Anzahl der nicht-matchenden Prefix-Komponenten gibt die Anzahl
			   * der Ebenen an, die wir aufsteigen können, weil es dort keine
			   * Treffer geben kann. 
			   * Wir gehen maximal 2 Ebenen höher, weil der
			   * Skriptbaum nicht alle Blätter auf der selben Höhe hat. 
			   * 2 Ebenen höher entspricht dem Übergang zur nächsten Library.
			   * Wir nehmen also an, dass innerhalb einer Library alle Skripts auf
			   * der selben Höhe sind, aber bei verschiedenen Libraries die
			   * Höhen unterschiedlich sein können.
			   */
			  int nMPC = nonMatchingPrefixComponents(prefixLC, nameToFindLC);
			  if (nMPC > 2) nMPC = 2;
			  if (nMPC > 0) return nMPC;
		
			  String nodeLocation = getLocation(node);
			  
			  if (location != null)
			  {
			    /* Falls die location des aktuellen Knotens nicht in der erlaubten
			     * Liste ist, können wir gleich 2 Ebenen aufsteigen (d.h zurück auf
			     * Ebene über Library), weil wir davon ausgehen können, dass innerhalb
			     * einer Library alle Skripte die selbe Location haben.
			     */
			    if (!stringInArray(nodeLocation, location)) return 2;
			  }
			  
			  //If the name doesn't even match case-insensitive, try the next sibling. 
			  if (!nameLC.equals(nameToFindLC[nameToFindLC.length - 1])) return 0;
			  
			  boolean isCaseCorrect = true;
			  prefix.add(name); //ACHTUNG! Muss nachher wieder entfernt werden
			  for (int i = nameToFind.length - 1, 
			           j = prefix.size() - 1;  i >= 0 && j>=0;  --i,--j)
			  {
			    if (!nameToFind[i].equals(prefix.get(j))) {isCaseCorrect = false; break;};
			  }
			  prefix.remove(prefix.size()-1); //wieder entfernen vor dem nächsten return
			  
			  FindNode findNode = new FindNode(node.unwrap(), xScriptProvider, nodeLocation, isCaseCorrect);
			  ListIterator liter = found.listIterator();
			  while (liter.hasNext())
			  {
			    if (findNode.betterMatchThan(liter.next(), location))
			    {liter.previous(); break;}
			  }
			  liter.add(findNode);
			  
			  /* ACHTUNG: Wir haben einen passenden Knoten gefunden. Nun könnten wir
			   * davon ausgehen, dass es im selben Modul keine weiteren Matches gibt
			   * und return 1 machen als Optimierung. Bei BASIC Makros ist dies
			   * auch korrekt, aber bei Makros in case-sensitiven Sprachen ist es
			   * durchaus möglich, dass im selben Modul mehrere Matches 
			   * (in unterschiedlicher Gross/Kleinschrift) sind.
			   * Mehr als return 1 ist auch bei BASIC nicht drin, weil
			   * auch BASIC bei Modul und Bibliotheksnamen case-sensitive ist.
		    */
			  if (getLanguage(node).equalsIgnoreCase("basic"))
			    return 1;
			  else
			    return 0;
			}
			else // if iter.hasNext()
			{
			  /* ACHTUNG! Diese Änderungen müssen vor return
			   * wieder Rückgängig gemacht werden
			   */
			  prefix.add(name);
	      prefixLC.add(name.toLowerCase());
			
	      while (iter.hasNext())
	      {
	        BrowseNode child = (BrowseNode)iter.next();
	        
	        int retL = findBrowseNodeTreeLeavesAndScriptProviders(child, prefix, prefixLC, nameToFind, nameToFindLC, location, xScriptProvider, found);
	        if (retL > 0) 
	        {
	          prefix.remove(prefix.size()-1);
	          prefixLC.remove(prefixLC.size()-1);
	          return retL - 1;
	        }
	      }
	      
	      prefix.remove(prefix.size()-1);
        prefixLC.remove(prefixLC.size()-1);
	      
			}
			return 0;
		}

	  /**
	   * Returns true iff array contains a String that is equals to str
     * @author bnk
     */
    private static boolean stringInArray(String str, String[] array)
    {
      for (int i = 0; i < array.length; ++i)
        if (str.equals(array[i])) return true;
      return false;
    }

    /** Falls <code>node.URL() == null</code> oder die URL keinen "location="
	   * Teil enthält, so wird "" geliefert, ansonsten der "location=" Teil ohne
	   * das führende "location=". 
     * @author bnk
     */
    private static String getLocation(BrowseNode node)
    { //T
	    return getUrlComponent(node, "location");
    }
    
    private static String getLanguage(BrowseNode node)
    {
	    return getUrlComponent(node, "language");
    }
    
    private static String getUrlComponent(BrowseNode node, String id)
    {
      String url = node.URL();
	    if (url == null) return "";
	    int idx = url.indexOf("?"+id+"=");
	    if (idx < 0) idx = url.indexOf("&"+id+"=");
	    if (idx < 0) return "";
	    idx += 10;
	    int idx2 = url.indexOf('&',idx);
	    if (idx2 < 0) idx2 = url.length();
      return url.substring(idx, idx2);
    }

    /**
     * Vergleicht von hinten beginnend die Strings in prefix mit den
     * Strings in nameToFind ohne sein letztes Element. Zurückgeliefert wird
     * die Länge der längsten Folge von dabei betrachteten Strings, die mit
     * einem gescheiterten Vergleich beginnt.
     * @author bnk
     */
    private static int nonMatchingPrefixComponents(List prefix, String[] nameToFind)
    { //T
      int i = nameToFind.length - 2; //beginne von hinten beim vorletzten Element
      int j = prefix.size()-1; //beginne von hinten beim letzten Element
      boolean matching = true;
      int count = 0;
      while (i >= 0 && j >= 0)
      {
        if (!matching || !nameToFind[i].equals(prefix.get(j)))
        {
          matching = false;
          ++count;
        }
        --i;
        --j;
      }
      return count;
    }

    /** 
     * Wenn <code>provider = null</code>, so wird versucht, 
     * einen passenden Provider zu finden.  
     * @author bnk
     */
    private static Object executeMacroInternal(String macroName, Object[] args, XScriptProvider provider, XBrowseNode root, String[] location)
    { //T
      XBrowseNodeAndXScriptProvider o = UNO.findBrowseNodeTreeLeafAndScriptProvider(root, "", macroName, false, location);
    	
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

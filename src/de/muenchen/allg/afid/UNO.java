/*
* Hilfsklasse zur leichteren Verwendung der UNO API.
* Dateiname: UNO.java
* Projekt  : n/a
* Funktion : Hilfsklasse zur leichteren Verwendung der UNO API.
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Datum     | Wer | Änderungsgrund
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
* ------------------------------------------------------------------- 
*
* @author D-III-ITD 5.1 Matthias S. Benkmann
* @version 1.0
* 
* */
package de.muenchen.allg.afid;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
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
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.XModifiable;
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
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner für Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link IllegalArgumentException} geworfen.
	 * @param args die Argumente, die dem Makro übergeben werden sollen.
	 * @throws RuntimeException wenn entweder kein passendes Makro gefunden wurde, oder
	 *              scriptProviderOrSupplier weder {@link XScriptProvider} noch
	 *              {@link XScriptProviderSupplier} implementiert.
	 * @return den Rückgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeMacro(Object scriptProviderOrSupplier, String macroName, Object[] args)
	{
		XScriptProvider provider = (XScriptProvider)UnoRuntime.queryInterface(XScriptProvider.class, scriptProviderOrSupplier);
		if (provider == null)
		{
			XScriptProviderSupplier supp = (XScriptProviderSupplier)UnoRuntime.queryInterface(XScriptProviderSupplier.class, scriptProviderOrSupplier);
			if (supp == null) throw new RuntimeException("Übergebenes Objekt ist weder XScriptProvider noch XScriptProviderSupplier");
			provider = supp.getScriptProvider();
		}
		
		XBrowseNode root = (XBrowseNode)UnoRuntime.queryInterface(XBrowseNode.class, provider);
		return Internal.executeMacroInternal(macroName, args, provider, root);
	}
	
	/**
	 * Ruft ein globales Makro auf (d,h, eines, das nicht in einem Dokument
	 * gespeichert ist).
	 * 
	 * @param macroName ist der Name des Makros. Der Name kann optional durch "." abgetrennte
	 *              Bezeichner für Bibliotheken/Module vorangestellt haben. Es sind also sowohl
	 *              "Foo" als auch "Module1.Foo" und "Standard.Module1.Foo" erlaubt.
	 *              Wenn kein passendes Makro gefunden wird, wird zuerst versucht, 
	 *              case-insensitive danach zu suchen. Falls dabei ebenfalls kein Makro
	 *              gefunden wird, wird eine {@link IllegalArgumentException} geworfen.
	 * @param args die Argumente, die dem Makro übergeben werden sollen.
	 * @throws RuntimeException wenn kein passendes Makro gefunden wurde.
	 * @return den Rückgabewert des Makros.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static Object executeGlobalMacro(String macroName, Object[] args)
	{
	  return Internal.executeMacroInternal(macroName, args, masterScriptProvider, scriptRoot.unwrap());
	}

		
	/**
	 * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt, dessen Name mit der
	 * Zeichefolge nameToFind endet. 
	 * @param xBrowseNode Wurzel des zu durchsuchenden Baums.
	 * @param prefix wird dem Namen jedes Knoten vorangestellt. Dies wird verwendet, wenn 
	 *          xBrowseNode nicht die Wurzel ist.
	 * @param nameToFind  das zu suchende Namenssuffix.
	 * @param caseSensitive falls true, so wird Gross-/Kleinschreibung berücksichtigt bei der 
	 *         Suche.
	 * @return den gefundenen Knoten oder null falls keiner gefunden.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode, String prefix, String nameToFind, boolean caseSensitive)
	{
		String name = xBrowseNode.getName();
		if (!prefix.equals(""))	name = prefix + "." + name; 
		
		if (!caseSensitive) 
		{ 
			name = name.toLowerCase(); 
			nameToFind = nameToFind.toLowerCase(); 
		}
		
		if (xBrowseNode.hasChildNodes())
		{
			XBrowseNode[] child = xBrowseNode.getChildNodes();
			for (int i = 0; i < child.length; ++i)
			{
				XBrowseNode o = findBrowseNodeTreeLeaf(child[i], name, nameToFind, caseSensitive);
				if (o != null) return o;
			}
		}
		else
			if (name.endsWith(nameToFind)) return xBrowseNode;
		
		return null;
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
		  return null;
		} catch (WrappedTargetException e) {
				e.printStackTrace();
				System.exit(0);
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

  /**
   * Interne Funktionen
   */
	private static class Internal
	{

    /**
     * @author bnk
     */
    private static Object executeMacroInternal(String macroName, Object[] args, XScriptProvider provider, XBrowseNode root)
    {
      Object oScript = UNO.findBrowseNodeTreeLeaf(root, "", "." + macroName, true);
    	if (oScript == null)
    		oScript = UNO.findBrowseNodeTreeLeaf(root, "", "." + macroName, false);
    	if (oScript == null)
    		oScript = UNO.findBrowseNodeTreeLeaf(root, "", macroName, true);
    	if (oScript == null)
    		oScript = UNO.findBrowseNodeTreeLeaf(root, "", macroName, false);
    	
    	XScript script;
    	
    	try{
    		XPropertySet props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, oScript);
    		String uri = (String)props.getPropertyValue("URI");
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

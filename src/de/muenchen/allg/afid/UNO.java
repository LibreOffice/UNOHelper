/*
* Hilfsklasse zur leichteren Verwendung der UNO API.
* Dateiname: UNO.java
* Projekt  : n/a
* Funktion : Hilfsklasse zur leichteren Verwendung der UNO API.
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Nr. |  Datum     |   Autor   | Änderungsgrund
* -------------------------------------------------------------------
* 001 | 26.04.2005 |    BNK    | Erstellung
* 002 | 07.07.2005 |    BNK    | Viele Verbesserungen
* 003 | 16.08.2005 |    BNK    | korrekte Dienststellenbezeichnung
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
import com.sun.star.script.browse.XBrowseNodeFactory;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.script.provider.XScriptProviderFactory;
import com.sun.star.sdb.XDocumentDataSource;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XURLTransformer;

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
	 * Initialisiert die statischen Felder dieser Klasse.
	 * @param remoteServiceManager der Haupt-ServiceManager.
	 * @throws Exception falls was schief geht.
	 * @author Matthias Benkmann (D-III-ITD 5.1)
	 */
	protected static void init(Object remoteServiceManager)
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

	/** Holt {@link XTextDocument} Interface von o.*/
	public static XTextDocument XTextDocument(Object o)
	{
		return (XTextDocument)UnoRuntime.queryInterface(XTextDocument.class,o);
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

}

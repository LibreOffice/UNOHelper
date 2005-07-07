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
* -------------------------------------------------------------------
*
* @author D-AfID 5.1 Matthias S. Benkmann
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
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.script.browse.BrowseNodeFactoryViewTypes;
import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.browse.XBrowseNodeFactory;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.script.provider.XScriptProviderFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

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
	 * The master script provider.
	 */
	public static XScriptProvider masterScriptProvider;
	
	/**
	 * Stellt die Verbindung mit OpenOffice her unter expliziter Angabe der
	 * Verbindungsparameter. Einfacher geht's mit {@link #init()}.
	 * @param connectionString z.B. "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager"
	 * @throws Exception falls was schief geht.
	 * @see init()
	 * @author Matthias Benkmann (D-HAIII 5.1)
	 */
	public static void init(String connectionString)
	throws Exception
	{
		XComponentContext xLocalContext = Bootstrap.createInitialComponentContext(null);
		XMultiComponentFactory xLocalFactory = xLocalContext.getServiceManager();
		XUnoUrlResolver xUrlResolver = (XUnoUrlResolver)UnoRuntime.queryInterface(XUnoUrlResolver.class,xLocalFactory.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",xLocalContext));
		init(xUrlResolver.resolve(connectionString));
	}
	
	/**
	 * Stellt die Verbindung mit OpenOffice her. Die Verbindungsparameter werden
	 * automagisch ermittelt.
	 * @throws Exception falls was schief geht.
	 * @author Matthias Benkmann (D-HAIII 5.1)
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
	 * @author Matthias Benkmann (D-HAIII 5.1)
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
	 * @author Matthias Benkmann (D-HAIII 5.1)
	 */
	public static XComponent loadComponentFromURL(String URL, 
																								boolean asTemplate, boolean allowMacros)
	throws com.sun.star.io.IOException, com.sun.star.lang.IllegalArgumentException
	{
		XComponentLoader loader = (XComponentLoader)UnoRuntime.queryInterface(XComponentLoader.class, desktop);
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
			XPropertySet props = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, o);
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
}

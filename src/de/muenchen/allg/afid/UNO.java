/*
* Hilfsklasse zur leichteren Verwendung der UNO API.
* Dateiname: UNO.java
* Projekt  : n/a
* Funktion : Hilfsklasse zur leichteren Verwendung der UNO API.
* 
 * Copyright (c) 2008 Landeshauptstadt München
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL),
 * version 1.0 (or any later version).
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * European Union Public Licence for more details.
 *
 * You should have received a copy of the European Union Public Licence
 * along with this program. If not, see
 * http://ec.europa.eu/idabc/en/document/7330
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
*                  | +XUpdatable()
*                  | +XTextViewCursorSupplier()
*                  | +XDrawPageSupplier()
* 16.06.2006 | BNK | +XContentEnumerationAccess()
* 19.06.2006 | LUT | +XControlShape()
* 20.06.2006 | LUT | +XTextRange()
* 29.06.2006 | LUT | +XDevice()
* 10.07.2006 | LUT | +XFramesSupplier()
* 28.07.2006 | BNK | +XInterface()
* 28.07.2006 | BNK | +XTextTable()
* 03.08.2006 | BNK | +XUserInputInterception()
* 04.08.2006 | LUT | +XDocumentInfoSupplier()
*                    +XDocumentInfo
* 07.08.2006 | BNK | +XDependentTextField
*                  | +XShape
*                  | +XTextFramesSupplier
* 10.08.2006 | BNK | +XTextFrame
* 21.08.2006 | BNK | +XSelectionSupplier
* 07.09.2006 | BNK | +XCell
* 13.09.2006 | LUT | +XDispatchProviderInterception
* 15.09.2006 | LUT | +XTextCursor
* 23.10.2006 | LUT | +XPageCursor
* 27.10.2006 | LUT | +XNotifyingDispatch
* 30.10.2006 | LUT | +getParsedUNOUrl()
* 30.10.2006 | BNK | +XDispatchHelper
* 30.10.2006 | LUT | +XStyleFamiliesSupplier()
* 30.10.2006 | LUT | +XStyle()
* 06.11.2006 | BNK | +dispatch(doc, url)
* 08.11.2006 | LUT | +getConfigurationAccess(nodepath)
*                    +getConfigurationUpdateAccess(nodepath)
* 01.12.2006 | LUT | +XTextRangeCompare
*                    +XTextSection
*                    +XTextSectionsSupplier
* 18.12.2006 | BNK | +XModuleUIConfigurationManagerSupplier
* 18.12.2006 | BAB | +XAcceleratorConfiguration
* 				   | +XUIConfigurationPersistence	
* 19.12.2006 | BAB | +getShortcutManager(component)
* 20.12.2006 | BNK | +XDataSource()
*                  | +XTablesSupplier()
*                  | +XColumnsSupplier()
* 21.12.2006 | BNK | +XKeysSupplier()
* 21.12.2006 | BNK | +XRow()
*                  | +XColumnLocate()
* 11.01.2007 | BNK | +XCellRangesQuery()
* 15.01.2007 | BNK | +loadComponentFromURL() dem man die Makro-Behandlung besser sagen kann
* 15.01.2007 | LUT | +XUIConfigurationManager
*                    +XModuleUIConfigurationManager
*                    +XIndexContainer()
* 29.01.2007 | LUT | +XFrame()
* 15.02.2007 | BAB | +XTextColumns()
* 22.02.2007 | BAB | +XStyleLoader()
* 24.04.2007 | LUT | +XPropertyState()
*                    +setPropertyToDefault(o, propName)
* 25.07.2007 | LUT | +XStringSubstitution()
* 19.09.2007 | BAB | +XRowSet()
* 16.10.2007 | BNK | +XCloseBroadcaster()
*                  | +XStorageBasedDocument()
* 06.11.2007 | BNK | +XTextGraphicObjectsSupplier()
* 14.01.2008 | BNK | +XTextPortionAppend()
* 16.01.2008 | BNK | +XTextContentAppend()
* 17.01.2008 | BNK | +XAutoTextContainer()
*                  | +XAutoTextGroup()
*                  | +XAutoTextEntry()
* 25.01.2008 | BNK | XTextContentAppend und XTextPortionAppend wieder entfernt
*                  | weil von alten Versionen nicht unterstützt.
* 08.07.2008 | LUT | +loadComponentFromURL mit Parameter hidden hinzugefügt.
* 11.07.2008 | BNK | +XFolderPicker()
* 04.12.2008 | BNK | +XController()
* 09.06.2009 | LUT | +dispatchAndWait(...)
* 25.06.2009 | LUT | +hideTextRange(...) zur korrekten Behandlung von Ein- 
*                  |  Ausblendungen
* 23.02.2010 | BNK | +XQueriesSupplier
* 04.05.2011 | LUT | +XDocumentMetadataAccess 
* 12.05.2011 | BED | +XRefreshable
* 13.05.2013 | UKT | Anpassungen an LO 4.0
* ------------------------------------------------------------------- 
*
* @author D-III-ITD 5.1 Matthias S. Benkmann
* 
* */
package de.muenchen.allg.afid;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.function.Consumer;

import com.sun.star.awt.XDevice;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XUserInputInterception;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindow2;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertyState;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XContentEnumerationAccess;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.document.XDocumentProperties;
import com.sun.star.document.XDocumentPropertiesSupplier;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.document.XStorageBasedDocument;
import com.sun.star.drawing.XControlShape;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.DispatchResultEvent;
import com.sun.star.frame.FrameSearchFlag;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XDispatchProviderInterception;
import com.sun.star.frame.XDispatchResultListener;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFramesSupplier;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XNotifyingDispatch;
import com.sun.star.frame.XStorable;
import com.sun.star.frame.XToolbarController;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.rdf.XDocumentMetadataAccess;
import com.sun.star.script.browse.BrowseNodeFactoryViewTypes;
import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.browse.XBrowseNodeFactory;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.script.provider.XScriptProviderFactory;
import com.sun.star.script.provider.XScriptProviderSupplier;
import com.sun.star.sdb.XDocumentDataSource;
import com.sun.star.sdb.XQueriesSupplier;
import com.sun.star.sdbc.XColumnLocate;
import com.sun.star.sdbc.XDataSource;
import com.sun.star.sdbc.XRow;
import com.sun.star.sdbc.XRowSet;
import com.sun.star.sdbcx.XColumnsSupplier;
import com.sun.star.sdbcx.XKeysSupplier;
import com.sun.star.sdbcx.XTablesSupplier;
import com.sun.star.sheet.XCellRangeData;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.style.XStyleLoader;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.text.XAutoTextContainer;
import com.sun.star.text.XAutoTextEntry;
import com.sun.star.text.XAutoTextGroup;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XDependentTextField;
import com.sun.star.text.XPageCursor;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextColumns;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XTextFrame;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextGraphicObjectsSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextRangeCompare;
import com.sun.star.text.XTextSection;
import com.sun.star.text.XTextSectionsSupplier;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.ui.XAcceleratorConfiguration;
import com.sun.star.ui.XModuleUIConfigurationManager;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIConfigurationPersistence;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.ui.dialogs.XFolderPicker;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.uno.XNamingService;
import com.sun.star.util.URL;
import com.sun.star.util.XChangesBatch;
import com.sun.star.util.XCloseBroadcaster;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XRefreshable;
import com.sun.star.util.XStringSubstitution;
import com.sun.star.util.XURLTransformer;
import com.sun.star.util.XUpdatable;
import com.sun.star.view.XPrintable;
import com.sun.star.view.XSelectionSupplier;
import com.sun.star.view.XViewSettingsSupplier;

/**
 * Hilfsklasse zur leichteren Verwendung der UNO API. * @author BNK
 *
 */
public class UNO
{
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
   * Ein com.sun.star.frame.XDispatchHelper.
   */
  public static XDispatchHelper dispatchHelper;
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
   * Enthält den configurationProvider, falls er bereits mit
   * getConfigurationProvider erzeugt wurde.
   */
  private static Object configurationProvider;

  /**
   * Stellt die Verbindung mit OpenOffice her unter expliziter Angabe der
   * Verbindungsparameter. Einfacher geht's mit {@link #init()}.
   * 
   * @param connectionString
   *          z.B.
   *          "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager"
   * @throws Exception
   *           falls was schief geht.
   * @see init()
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static void init(String connectionString) throws Exception
  {
    XComponentContext xLocalContext = Bootstrap
        .createInitialComponentContext(null);
    XMultiComponentFactory xLocalFactory = xLocalContext.getServiceManager();
    XUnoUrlResolver xUrlResolver = UNO
        .XUnoUrlResolver(xLocalFactory.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", xLocalContext));
    init(xUrlResolver.resolve(connectionString));
  }

  /**
   * Stellt die Verbindung mit OpenOffice her. Die Verbindungsparameter werden
   * automagisch ermittelt.
   * 
   * @throws Exception
   *           falls was schief geht.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static void init() throws Exception
  {
    init(Bootstrap.bootstrap().getServiceManager());
  }

  /**
   * Initialisiert die statischen Felder dieser Klasse ausgehend für eine
   * existierende Verbindung.
   * 
   * @param remoteServiceManager
   *          der Haupt-ServiceManager.
   * @throws Exception
   *           falls was schief geht.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static void init(Object remoteServiceManager) throws Exception
  {
    xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
        remoteServiceManager);
    xMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xMCF);
    defaultContext = UnoRuntime.queryInterface(XComponentContext.class,
        (UnoRuntime.queryInterface(XPropertySet.class, xMCF))
            .getPropertyValue("DefaultContext"));

    desktop = UnoRuntime.queryInterface(XDesktop.class,
        xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",
            defaultContext));

    XBrowseNodeFactory masterBrowseNodeFac = UnoRuntime
        .queryInterface(XBrowseNodeFactory.class, defaultContext.getValueByName(
            "/singletons/com.sun.star.script.browse.theBrowseNodeFactory"));
    scriptRoot = new BrowseNode(masterBrowseNodeFac
        .createView(BrowseNodeFactoryViewTypes.MACROORGANIZER));

    XScriptProviderFactory masterProviderFac = UnoRuntime.queryInterface(
        XScriptProviderFactory.class, defaultContext.getValueByName(
            "/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"));
    masterScriptProvider = masterProviderFac
        .createScriptProvider(defaultContext);

    dbContext = UNO.XNamingService(
        UNO.createUNOService("com.sun.star.sdb.DatabaseContext"));
    urlTransformer = UNO.XURLTransformer(
        UNO.createUNOService("com.sun.star.util.URLTransformer"));
    dispatchHelper = UNO.XDispatchHelper(
        UNO.createUNOService("com.sun.star.frame.DispatchHelper"));
  }

  /**
   * Läd ein Dokument und setzt im Erfolgsfall {@link #compo} auf das geöffnete
   * Dokument.
   * 
   * @param url
   *          die URL des zu ladenden Dokuments, z.B.
   *          "file:///C:/temp/footest.odt" oder "private:factory/swriter" (für
   *          ein leeres).
   * @param asTemplate
   *          falls true wird das Dokument als Vorlage behandelt und ein neues
   *          unbenanntes Dokument erzeugt.
   * @param allowMacros
   *          falls true wird die Ausführung von Makros freigeschaltet.
   * @return das geöffnete Dokument
   * @throws com.sun.star.io.IOException
   * @throws com.sun.star.lang.IllegalArgumentException
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
      boolean allowMacros) throws com.sun.star.io.IOException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Läd ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url
   *          die URL des zu ladenden Dokuments, z.B.
   *          "file:///C:/temp/footest.odt" oder "private:factory/swriter" (für
   *          ein leeres).
   * @param asTemplate
   *          falls true wird das Dokument als Vorlage behandelt und ein neues
   *          unbenanntes Dokument erzeugt.
   * @param allowMacros
   *          falls true wird die Ausführung von Makros freigeschaltet.
   * @param hidden
   *          falls true wird das Dokument unsichtbar geöffnet
   * @return das geöffnete Dokument
   * @throws com.sun.star.io.IOException
   * @throws com.sun.star.lang.IllegalArgumentException
   * @author Christoph Lutz (D-III-ITD 5.1)
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
      boolean allowMacros, boolean hidden) throws com.sun.star.io.IOException
  {
    short allowMacrosShort;
    if (allowMacros)
      allowMacrosShort = MacroExecMode.ALWAYS_EXECUTE_NO_WARN;
    else
      allowMacrosShort = MacroExecMode.NEVER_EXECUTE;
    return loadComponentFromURL(url, asTemplate, allowMacrosShort, hidden);
  }

  /**
   * Läd ein Dokument und setzt im Erfolgsfall {@link #compo} auf das geöffnete
   * Dokument.
   * 
   * @param url
   *          die URL des zu ladenden Dokuments, z.B.
   *          "file:///C:/temp/footest.odt" oder "private:factory/swriter" (für
   *          ein leeres).
   * @param asTemplate
   *          falls true wird das Dokument als Vorlage behandelt und ein neues
   *          unbenanntes Dokument erzeugt.
   * @param allowMacros
   *          eine der Konstanten aus {@link MacroExecMode}.
   * @return das geöffnete Dokument
   * @throws com.sun.star.io.IOException
   * @throws com.sun.star.lang.IllegalArgumentException
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
      short allowMacros) throws com.sun.star.io.IOException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Läd ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url
   *          die URL des zu ladenden Dokuments, z.B.
   *          "file:///C:/temp/footest.odt" oder "private:factory/swriter" (für
   *          ein leeres).
   * @param asTemplate
   *          falls true wird das Dokument als Vorlage behandelt und ein neues
   *          unbenanntes Dokument erzeugt.
   * @param allowMacros
   *          eine der Konstanten aus {@link MacroExecMode}.
   * @return das geöffnete Dokument
   * @throws com.sun.star.io.IOException
   * @throws com.sun.star.lang.IllegalArgumentException
   * @author Christoph Lutz (D-III-ITD 5.1)
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
      short allowMacros, boolean hidden) throws com.sun.star.io.IOException
  {
    XComponentLoader loader = UNO.XComponentLoader(desktop);
    PropertyValue[] arguments = new PropertyValue[3];
    arguments[0] = new PropertyValue();
    arguments[0].Name = "MacroExecutionMode";
    arguments[0].Value = Short.valueOf(allowMacros);
    arguments[1] = new PropertyValue();
    arguments[1].Name = "AsTemplate";
    arguments[1].Value = Boolean.valueOf(asTemplate);
    arguments[2] = new PropertyValue();
    arguments[2].Name = "Hidden";
    arguments[2].Value = Boolean.valueOf(hidden);
    XComponent lc = loader.loadComponentFromURL(url, "_blank",
        FrameSearchFlag.CREATE, arguments);
    if (lc != null)
      compo = lc;
    return lc;
  }

  /**
   * Dispatcht url auf dem aktuellen Controll von doc.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static void dispatch(XModel doc, String url)
  {
    XDispatchProvider prov = UNO
        .XDispatchProvider(doc.getCurrentController().getFrame());
    dispatchHelper.executeDispatch(prov, url, "", FrameSearchFlag.SELF,
        new PropertyValue[] {});
  }

  /**
   * Dispatcht url auf dem aktuellen Controll von doc mittels
   * {@link XNotifyingDispatch}, wartet auf die Benachrichtigung, dass der
   * Dispatch vollständig abgearbeitet ist, und liefert das
   * {@link DispatchResultEvent} dieser Benachrichtigung zurück oder null, wenn
   * zu url kein XNotifyingDispatch definiert ist oder der Dispatch mit
   * disposing abgebrochen wurde.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static DispatchResultEvent dispatchAndWait(XTextDocument doc,
      String url)
  {
    if (doc == null)
      return null;

    URL unoUrl = getParsedUNOUrl(url);

    XDispatchProvider prov = UNO
        .XDispatchProvider(doc.getCurrentController().getFrame());
    if (prov == null)
      return null;

    XNotifyingDispatch disp = UNO.XNotifyingDispatch(
        prov.queryDispatch(unoUrl, "", FrameSearchFlag.SELF));
    if (disp == null)
      return null;

    final boolean[] lock = new boolean[] { true };
    final DispatchResultEvent[] resultEvent = new DispatchResultEvent[] {
        null };

    disp.dispatchWithNotification(unoUrl, new PropertyValue[] {},
        new XDispatchResultListener()
        {
          @Override
          public void disposing(EventObject arg0)
          {
            synchronized (lock)
            {
              lock[0] = false;
              lock.notifyAll();
            }
          }

          @Override
          public void dispatchFinished(DispatchResultEvent arg0)
          {
            synchronized (lock)
            {
              resultEvent[0] = arg0;
              lock[0] = false;
              lock.notifyAll();
            }
          }
        });

    synchronized (lock)
    {
      while (lock[0])
        try
        {
          lock.wait();
        } catch (InterruptedException e)
        {
        }
    }
    return resultEvent[0];
  }

  /**
   * Ruft ein Makro auf unter expliziter Angabe der Komponente, die es zur
   * Verfügung stellt.
   * 
   * @param scriptProviderOrSupplier
   *          ist ein Objekt, das entweder {@link XScriptProvider} oder
   *          {@link XScriptProviderSupplier} implementiert. Dies kann z.B. ein
   *          TextDocument sein. Soll einfach nur ein Skript aus dem gesamten
   *          Skript-Baum ausgeführt werden, kann die Funktion
   *          {@link #executeGlobalMacro(String, Object[])} verwendet werden,
   *          die diesen Parameter nicht erfordert. ACHTUNG! Es wird nicht
   *          zwangsweise der übergebene scriptProviderOrSupplier verwendet um
   *          das Skript auszuführen. Er stellt nur den Einstieg in den
   *          Skript-Baum dar.
   * @param macroName
   *          ist der Name des Makros. Der Name kann optional durch "."
   *          abgetrennte Bezeichner für Bibliotheken/Module vorangestellt
   *          haben. Es sind also sowohl "Foo" als auch "Module1.Foo" und
   *          "Standard.Module1.Foo" erlaubt. Wenn kein passendes Makro gefunden
   *          wird, wird zuerst versucht, case-insensitive danach zu suchen.
   *          Falls dabei ebenfalls kein Makro gefunden wird, wird eine
   *          {@link RuntimeException} geworfen.
   * @param args
   *          die Argumente, die dem Makro übergeben werden sollen.
   * @param location
   *          eine Liste aller erlaubten locations ("application", "share",
   *          "document") für das Makro. Bei der Suche wird zuerst ein
   *          case-sensitive Match in allen gelisteten locations gesucht, bevor
   *          die case-insensitive Suche versucht wird. Durch Verwendung der
   *          exakten Gross-/Kleinschreibung des Makros und korrekte Ordnung der
   *          location Liste lässt sich also immer das richtige Makro
   *          selektieren.
   * @throws RuntimeException
   *           wenn entweder kein passendes Makro gefunden wurde, oder
   *           scriptProviderOrSupplier weder {@link XScriptProvider} noch
   *           {@link XScriptProviderSupplier} implementiert.
   * @return den Rückgabewert des Makros.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static Object executeMacro(Object scriptProviderOrSupplier,
      String macroName, Object[] args, String[] location)
  {
    XScriptProvider provider = UnoRuntime.queryInterface(XScriptProvider.class,
        scriptProviderOrSupplier);
    if (provider == null)
    {
      XScriptProviderSupplier supp = UnoRuntime.queryInterface(
          XScriptProviderSupplier.class, scriptProviderOrSupplier);
      if (supp == null)
        throw new RuntimeException(
            "Übergebenes Objekt ist weder XScriptProvider noch XScriptProviderSupplier");
      provider = supp.getScriptProvider();
    }

    XBrowseNode root = UnoRuntime.queryInterface(XBrowseNode.class, provider);
    /*
     * Wir übergeben NICHT provider als drittes Argument, sondern lassen
     * Internal.executeMacroInternal den provider selbst bestimmen. Das hat
     * keinen besonderen Grund. Es erscheint einfach nur etwas robuster, den
     * "nächstgelegenen" ScriptProvider zu verwenden.
     */
    return Utils.executeMacroInternal(macroName, args, null, root, location);
  }

  /**
   * Ruft ein globales Makro auf (d,h, eines, das nicht in einem Dokument
   * gespeichert ist). Im Falle gleichnamiger Makros hat ein Makro mit
   * location=application Vorrang vor einem mit location=share.
   * 
   * @param macroName
   *          ist der Name des Makros. Der Name kann optional durch "."
   *          abgetrennte Bezeichner für Bibliotheken/Module vorangestellt
   *          haben. Es sind also sowohl "Foo" als auch "Module1.Foo" und
   *          "Standard.Module1.Foo" erlaubt. Wenn kein passendes Makro gefunden
   *          wird, wird zuerst versucht, case-insensitive danach zu suchen.
   *          Falls dabei ebenfalls kein Makro gefunden wird, wird eine
   *          {@link RuntimeException} geworfen.
   * @param args
   *          die Argumente, die dem Makro übergeben werden sollen.
   * @throws RuntimeException
   *           wenn kein passendes Makro gefunden wurde.
   * @return den Rückgabewert des Makros.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static Object executeGlobalMacro(String macroName, Object[] args)
  {
    final String[] userAndShare = new String[] { "application", "share" };
    return Utils.executeMacroInternal(macroName, args, null,
        scriptRoot.unwrap(), userAndShare);
  }

  /**
   * Ruft ein Makro aus dem gesamten Makro-Baum auf.
   * 
   * @param macroName
   *          ist der Name des Makros. Der Name kann optional durch "."
   *          abgetrennte Bezeichner für Bibliotheken/Module vorangestellt
   *          haben. Es sind also sowohl "Foo" als auch "Module1.Foo" und
   *          "Standard.Module1.Foo" erlaubt. Wenn kein passendes Makro gefunden
   *          wird, wird zuerst versucht, case-insensitive danach zu suchen.
   *          Falls dabei ebenfalls kein Makro gefunden wird, wird eine
   *          {@link RuntimeException} geworfen.
   * @param args
   *          die Argumente, die dem Makro übergeben werden sollen.
   * @param location
   *          eine Liste aller erlaubten locations ("application", "share",
   *          "document") für das Makro. Bei der Suche wird zuerst ein
   *          case-sensitive Match in allen gelisteten locations gesucht, bevor
   *          die case-insenstive Suche versucht wird. Durch Verwendung der
   *          exakten Gross-/Kleinschreibung des Makros und korrekte Ordnung der
   *          location Liste lässt sich also immer das richtige Makro
   *          selektieren.
   * @throws RuntimeException
   *           wenn kein passendes Makro gefunden wurde.
   * @return den Rückgabewert des Makros.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static Object executeMacro(String macroName, Object[] args,
      String[] location)
  {
    return Utils.executeMacroInternal(macroName, args, null,
        scriptRoot.unwrap(), location);
  }

  /**
   * Diese Methode prüft, ob es sich bei dem übergebenen Objekt service um ein
   * UNO-Service mit dem Namen serviceName handelt und liefert true zurück, wenn
   * das Objekt das XServiceInfo-Interface und den gesuchten Service
   * implementiert, ansonsten wird false zurückgegeben.
   * 
   * @param service
   *          Das zu prüfende Service-Objekt
   * @param serviceName
   *          der voll-qualifizierte Service-Name des services.
   * @return true, wenn das Objekt das XServiceInfo-Interface und den gesuchten
   *         Service implementiert, ansonsten false.
   * @author Christoph Lutz (D-III-ITD 5.1)
   */
  public static boolean supportsService(Object service, String serviceName)
  {
    if (UNO.XServiceInfo(service) != null)
      return UNO.XServiceInfo(service).supportsService(serviceName);
    return false;
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT,
   * dessen Name nameToFind ist (kann durch "." abgetrennte Pfadangabe im
   * Skript-Baum enthalten). Siehe
   * {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}
   * .
   * 
   * @return den gefundenen Knoten oder null falls keiner gefunden.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode,
      String prefix, String nameToFind, boolean caseSensitive)
  {
    XBrowseNodeAndXScriptProvider x = findBrowseNodeTreeLeafAndScriptProvider(
        xBrowseNode, prefix, nameToFind, caseSensitive);
    return x.XBrowseNode;
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT,
   * dessen Name nameToFind ist (kann durch "." abgetrennte Pfadangabe im
   * Skript-Baum enthalten). Siehe
   * {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}
   * .
   * 
   * @return den gefundenen Knoten, sowie den nächsten Vorfahren, der
   *         XScriptProvider implementiert (oder den Knoten selbst, falls dieser
   *         XScriptProvider implementiert). Falls kein entsprechender Knoten
   *         oder Vorfahre gefunden wurde, wird der entsprechende Wert als null
   *         geliefert.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(
      XBrowseNode xBrowseNode, String prefix, String nameToFind,
      boolean caseSensitive)
  {
    return findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix,
        nameToFind, caseSensitive, null);
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT,
   * dessen Name nameToFind ist (kann durch "." abgetrennte Pfadangabe im
   * Skript-Baum enthalten). Siehe
   * {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String)}
   * .
   * 
   * @param xBrowseNode
   *          Wurzel des zu durchsuchenden Baums.
   * @param prefix
   *          wird dem Namen jedes Knoten vorangestellt. Dies wird verwendet,
   *          wenn xBrowseNode nicht die Wurzel ist.
   * @param nameToFind
   *          der zu suchende Name.
   * @param caseSensitive
   *          falls true, so wird Gross-/Kleinschreibung berücksichtigt bei der
   *          Suche.
   * @param location
   *          Es gelten nur Knoten als Treffer, die ein "URI" Property haben,
   *          das eine location enthält die einem String in der
   *          <code>location</code> Liste entspricht. Mögliche locations sind
   *          "document", "application" und "share". Falls
   *          <code>location==null</code>, so wird {"document", "application",
   *          "share"} angenommen.
   * @return den gefundenen Knoten, sowie den nächsten Vorfahren, der
   *         XScriptProvider implementiert. Falls kein entsprechender Knoten
   *         oder Vorfahre gefunden wurde, wird der entsprechende Wert als null
   *         geliefert.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(
      XBrowseNode xBrowseNode, String prefix, String nameToFind,
      boolean caseSensitive, String[] location)
  {
    String[] loc;
    if (location == null)
    {
      loc = new String[] { "document", "application", "share" };
    } else
    {
      loc = location;
    }
    List<Utils.FindNode> found = new LinkedList<Utils.FindNode>();
    List<String> prefixVec = new ArrayList<String>();
    List<String> prefixLCVec = new ArrayList<String>();
    String[] prefixArr = prefix.split("\\.");
    for (int i = 0; i < prefixArr.length; ++i)
      if (!prefixArr[i].isEmpty())
      {
        prefixVec.add(prefixArr[i]);
        prefixLCVec.add(prefixArr[i].toLowerCase());
      }

    String[] nameToFindArrPre = nameToFind.split("\\.");
    int i1 = 0;
    while (i1 < nameToFindArrPre.length && nameToFindArrPre[i1].isEmpty())
      ++i1;
    int i2 = nameToFindArrPre.length - 1;
    while (i2 >= i1 && nameToFindArrPre[i2].isEmpty())
      --i2;
    ++i2;
    String[] nameToFindArr = new String[i2 - i1];
    String[] nameToFindLCArr = new String[i2 - i1];
    for (int i = 0; i < nameToFindArr.length; ++i)
    {
      nameToFindArr[i] = nameToFindArrPre[i1];
      nameToFindLCArr[i] = nameToFindArrPre[i1].toLowerCase();
      ++i1;
    }

    Utils.findBrowseNodeTreeLeavesAndScriptProviders(
        new BrowseNode(xBrowseNode), prefixVec, prefixLCVec, nameToFindArr,
        nameToFindLCArr, loc, null, found);

    if (found.isEmpty())
      return new XBrowseNodeAndXScriptProvider(null, null);

    Utils.FindNode findNode = found.get(0);

    if (caseSensitive && !findNode.isCaseCorrect)
      return new XBrowseNodeAndXScriptProvider(null, null);

    return new XBrowseNodeAndXScriptProvider(findNode.XBrowseNode,
        findNode.XScriptProvider);
  }

  /**
   * Liefert den Wert von Property propName des Objekts o zurück.
   * 
   * @return den Wert des Propertys oder <code>null</code>, falls o entweder
   *         nicht das XPropertySet Interface implementiert, oder kein Property
   *         names propName hat oder ein sonstiger Fehler aufgetreten ist.
   */
  public static Object getProperty(Object o, String propName)
  {
    Object ret = null;
    try
    {
      XPropertySet props = UNO.XPropertySet(o);
      if (props == null)
        return null;
      ret = props.getPropertyValue(propName);
    } catch (Exception e)
    {
    }
    return ret;
  }

  /**
   * Setzt das Property propName des Objekts o auf den Wert propVal und liefert
   * den neuen Wert zurück. Falls o kein XPropertySet implementiert, oder das
   * Property propName nicht gelesen werden kann (z.B. weil o diese Property
   * nicht besitzt), so wird null zurückgeliefert. Zu beachten ist, dass es
   * möglich ist, dass der zurückgelieferte Wert nicht propVal und auch nicht
   * null ist. Dies geschieht insbesondere, wenn ein Event Handler sein Veto
   * gegen die Änderung eingelegt hat.
   * 
   * @param o
   *          das Objekt, dessen Property zu ändern ist.
   * @param propName
   *          der Name des zu ändernden Properties.
   * @param propVal
   *          der neue Wert.
   * @return der Wert des Propertys nach der (versuchten) Änderung oder null,
   *         falls der Wert des Propertys nicht mal lesbar ist.
   * @author bnk
   */
  public static Object setProperty(Object o, String propName, Object propVal)
  {
    Object ret = null;
    try
    {
      XPropertySet props = UNO.XPropertySet(o);
      if (props == null)
        return null;
      try
      {
        props.setPropertyValue(propName, propVal);
      } catch (Exception x)
      {
      }
      ret = props.getPropertyValue(propName);
    } catch (Exception e)
    {
    }
    return ret;
  }

  /**
   * Setzt das Property propName auf den ursprünglichen Wert zurück, der als
   * Voreinstellung für das Objekt o hinterlegt ist und liefert den neuen Wert
   * zurück. Falls o kein XPropertyState implementiert, oder das Property
   * propName nicht gelesen werden kann (z.B. weil o diese Property nicht
   * besitzt), so wird null zurückgeliefert. Zu beachten ist, dass es möglich
   * ist, dass das Property nicht zurück gesetzt wird, wenn ein Event Handler
   * sein Veto gegen die Änderung einlegt.
   * 
   * @param o
   *          das Objekt, dessen Property zu ändern ist.
   * @param propName
   *          der Name des zu ändernden Properties.
   * @param propVal
   *          der neue Wert.
   * @return der Wert des Propertys nach der (versuchten) Änderung oder null,
   *         falls der Wert des Propertys nicht mal lesbar ist.
   * @author bnk
   */
  public static Object setPropertyToDefault(Object o, String propName)
  {
    Object ret = null;
    try
    {
      XPropertyState props = UNO.XPropertyState(o);
      if (props == null)
        return null;
      try
      {
        props.setPropertyToDefault(propName);
      } catch (Exception x)
      {
      }
      ret = getProperty(props, propName);
    } catch (Exception e)
    {
    }
    return ret;
  }

  /**
   * Erzeugt einen Dienst im Haupt-Servicemanager mit dem DefaultContext.
   * 
   * @param serviceName
   *          Name des zu erzeugenden Dienstes
   * @return ein Objekt, das den Dienst anbietet, oder null falls Fehler.
   * @author bnk
   */
  public static Object createUNOService(String serviceName)
  {
    try
    {
      return xMCF.createInstanceWithContext(serviceName, defaultContext);
    } catch (Exception x)
    {
      return null;
    }
  }
  
  /** Holt {@link XSingleServiceFactory} Interface von o. */
  public static XSingleServiceFactory XSingleServiceFactory(Object o)
  {
    return UnoRuntime.queryInterface(XSingleServiceFactory.class, o);
  }

  /** Holt {@link XViewSettingsSupplier} Interface von o. */
  public static XViewSettingsSupplier XViewSettingsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XViewSettingsSupplier.class, o);
  }

  /** Holt {@link XStorable} Interface von o. */
  public static XStorable XStorable(Object o)
  {
    return UnoRuntime.queryInterface(XStorable.class, o);
  }

  /** Holt {@link XAutoTextContainer} Interface von o. */
  public static XAutoTextContainer XAutoTextContainer(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextContainer.class, o);
  }

  /** Holt {@link XAutoTextGroup} Interface von o. */
  public static XAutoTextGroup XAutoTextGroup(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextGroup.class, o);
  }

  /** Holt {@link XAutoTextEntry} Interface von o. */
  public static XAutoTextEntry XAutoTextEntry(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextEntry.class, o);
  }

  /** Holt {@link XShape} Interface von o. */
  public static XShape XShape(Object o)
  {
    return UnoRuntime.queryInterface(XShape.class, o);
  }

  /** Holt {@link XTopWindow} Interface von o. */
  public static XTopWindow XTopWindow(Object o)
  {
    return UnoRuntime.queryInterface(XTopWindow.class, o);
  }

  /** Holt {@link XCloseable} Interface von o. */
  public static XCloseable XCloseable(Object o)
  {
    return UnoRuntime.queryInterface(XCloseable.class, o);
  }

  /** Holt {@link XCloseBroadcaster} Interface von o. */
  public static XCloseBroadcaster XCloseBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XCloseBroadcaster.class, o);
  }

  /** Holt {@link XStorageBasedDocument} Interface von o. */
  public static XStorageBasedDocument XStorageBasedDocument(Object o)
  {
    return UnoRuntime.queryInterface(XStorageBasedDocument.class, o);
  }

  /** Holt {@link XWindow2} Interface von o. */
  public static XWindow2 XWindow2(Object o)
  {
    return UnoRuntime.queryInterface(XWindow2.class, o);
  }

  /** Holt {@link XTextFramesSupplier} Interface von o. */
  public static XTextFramesSupplier XTextFramesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextFramesSupplier.class, o);
  }

  /** Holt {@link XTextGraphicObjectsSupplier} Interface von o. */
  public static XTextGraphicObjectsSupplier XTextGraphicObjectsSupplier(
      Object o)
  {
    return UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, o);
  }

  /** Holt {@link XTextFrame} Interface von o. */
  public static XTextFrame XTextFrame(Object o)
  {
    return UnoRuntime.queryInterface(XTextFrame.class, o);
  }

  /** Holt {@link XTextFieldsSupplier} Interface von o. */
  public static XTextFieldsSupplier XTextFieldsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextFieldsSupplier.class, o);
  }

  /** Holt {@link XComponent} Interface von o. */
  public static XComponent XComponent(Object o)
  {
    return UnoRuntime.queryInterface(XComponent.class, o);
  }

  /** Holt {@link XEventBroadcaster} Interface von o. */
  public static XEventBroadcaster XEventBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XEventBroadcaster.class, o);
  }

  /** Holt {@link XBrowseNode} Interface von o. */
  public static XBrowseNode XBrowseNode(Object o)
  {
    return UnoRuntime.queryInterface(XBrowseNode.class, o);
  }

  /** Holt {@link XCellRangeData} Interface von o. */
  public static XCellRangeData XCellRangeData(Object o)
  {
    return UnoRuntime.queryInterface(XCellRangeData.class, o);
  }

  /** Holt {@link XCell} Interface von o. */
  public static XCell XCell(Object o)
  {
    return UnoRuntime.queryInterface(XCell.class, o);
  }

  /** Holt {@link XColumnsSupplier} Interface von o. */
  public static XColumnsSupplier XColumnsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XColumnsSupplier.class, o);
  }

  /** Holt {@link XTableColumns} Interface von o. */
  public static XTableColumns XTableColumns(Object o)
  {
    return UnoRuntime.queryInterface(XTableColumns.class, o);
  }

  /** Holt {@link XColumnRowRange} Interface von o. */
  public static XColumnRowRange XColumnRowRange(Object o)
  {
    return UnoRuntime.queryInterface(XColumnRowRange.class, o);
  }

  /** Holt {@link XIndexAccess} Interface von o. */
  public static XIndexAccess XIndexAccess(Object o)
  {
    return UnoRuntime.queryInterface(XIndexAccess.class, o);
  }

  /** Holt {@link XCellRange} Interface von o. */
  public static XCellRange XCellRange(Object o)
  {
    return UnoRuntime.queryInterface(XCellRange.class, o);
  }

  /** Holt {@link XUserInputInterception} Interface von o. */
  public static XUserInputInterception XUserInputInterception(Object o)
  {
    return UnoRuntime.queryInterface(XUserInputInterception.class, o);
  }

  /** Holt {@link XModifiable} Interface von o. */
  public static XModifiable XModifiable(Object o)
  {
    return UnoRuntime.queryInterface(XModifiable.class, o);
  }

  /** Holt {@link XRow} Interface von o. */
  public static XRow XRow(Object o)
  {
    return UnoRuntime.queryInterface(XRow.class, o);
  }

  /** Holt {@link XCellRangesQuery} Interface von o. */
  public static XCellRangesQuery XCellRangesQuery(Object o)
  {
    return UnoRuntime.queryInterface(XCellRangesQuery.class, o);
  }

  /** Holt {@link XDocumentDataSource} Interface von o. */
  public static XDocumentDataSource XDocumentDataSource(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentDataSource.class, o);
  }

  /** Holt {@link XDataSource} Interface von o. */
  public static XDataSource XDataSource(Object o)
  {
    return UnoRuntime.queryInterface(XDataSource.class, o);
  }

  /** Holt {@link XTablesSupplier} Interface von o. */
  public static XTablesSupplier XTablesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTablesSupplier.class, o);
  }

  /** Holt {@link XQueriesSupplier} Interface von o. */
  public static XQueriesSupplier XQueriesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XQueriesSupplier.class, o);
  }

  /** Holt {@link XNamingService} Interface von o. */
  public static XNamingService XNamingService(Object o)
  {
    return UnoRuntime.queryInterface(XNamingService.class, o);
  }

  /** Holt {@link XNamed} Interface von o. */
  public static XNamed XNamed(Object o)
  {
    return UnoRuntime.queryInterface(XNamed.class, o);
  }

  /** Holt {@link XUnoUrlResolver} Interface von o. */
  public static XUnoUrlResolver XUnoUrlResolver(Object o)
  {
    return UnoRuntime.queryInterface(XUnoUrlResolver.class, o);
  }

  /** Holt {@link XMultiPropertySet} Interface von o. */
  public static XMultiPropertySet XMultiPropertySet(Object o)
  {
    return UnoRuntime.queryInterface(XMultiPropertySet.class, o);
  }

  /** Holt {@link XPropertySet} Interface von o. */
  public static XPropertySet XPropertySet(Object o)
  {
    return UnoRuntime.queryInterface(XPropertySet.class, o);
  }

  /** Holt {@link XModel} Interface von o. */
  public static XModel XModel(Object o)
  {
    return UnoRuntime.queryInterface(XModel.class, o);
  }

  /** Holt {@link XController} Interface von o. */
  public static XController XController(Object o)
  {
    return UnoRuntime.queryInterface(XController.class, o);
  }

  /** Holt {@link XTextField} Interface von o. */
  public static XTextField XTextField(Object o)
  {
    return UnoRuntime.queryInterface(XTextField.class, o);
  }

  /** Holt {@link XDocumentInsertable} Interface von o. */
  public static XDocumentInsertable XDocumentInsertable(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentInsertable.class, o);
  }

  /** Holt {@link XModifyBroadcaster} Interface von o. */
  public static XModifyBroadcaster XModifyBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XModifyBroadcaster.class, o);
  }

  /** Holt {@link XScriptProvider} Interface von o. */
  public static XScriptProvider XScriptProvider(Object o)
  {
    return UnoRuntime.queryInterface(XScriptProvider.class, o);
  }

  /** Holt {@link XSpreadsheet} Interface von o. */
  public static XSpreadsheet XSpreadsheet(Object o)
  {
    return UnoRuntime.queryInterface(XSpreadsheet.class, o);
  }

  /** Holt {@link XModuleUIConfigurationManagerSupplier} Interface von o. */
  public static XModuleUIConfigurationManagerSupplier XModuleUIConfigurationManagerSupplier(
      Object o)
  {
    return UnoRuntime
        .queryInterface(XModuleUIConfigurationManagerSupplier.class, o);
  }

  /** Holt {@link XText} Interface von o. */
  public static XText XText(Object o)
  {
    return UnoRuntime.queryInterface(XText.class, o);
  }

  /** Holt {@link XTextTable} Interface von o. */
  public static XTextTable XTextTable(Object o)
  {
    return UnoRuntime.queryInterface(XTextTable.class, o);
  }

  /** Holt {@link XDependentTextField} Interface von o. */
  public static XDependentTextField XDependentTextField(Object o)
  {
    return UnoRuntime.queryInterface(XDependentTextField.class, o);
  }

  /** Holt {@link XLayoutManager} Interface von o. */
  public static XLayoutManager XLayoutManager(Object o)
  {
    return UnoRuntime.queryInterface(XLayoutManager.class, o);
  }

  /** Holt {@link XColumnLocate} Interface von o. */
  public static XColumnLocate XColumnLocate(Object o)
  {
    return UnoRuntime.queryInterface(XColumnLocate.class, o);
  }

  /** Holt {@link XEnumerationAccess} Interface von o. */
  public static XEnumerationAccess XEnumerationAccess(Object o)
  {
    return UnoRuntime.queryInterface(XEnumerationAccess.class, o);
  }

  /** Holt {@link XSpreadsheetDocument} Interface von o. */
  public static XSpreadsheetDocument XSpreadsheetDocument(Object o)
  {
    return UnoRuntime.queryInterface(XSpreadsheetDocument.class, o);
  }

  /** Holt {@link XPrintable} Interface von o. */
  public static XPrintable XPrintable(Object o)
  {
    return UnoRuntime.queryInterface(XPrintable.class, o);
  }

  /** Holt {@link XTextDocument} Interface von o. */
  public static XTextDocument XTextDocument(Object o)
  {
    return UnoRuntime.queryInterface(XTextDocument.class, o);
  }

  /** Holt {@link XTextContent} Interface von o. */
  public static XTextContent XTextContent(Object o)
  {
    return UnoRuntime.queryInterface(XTextContent.class, o);
  }

  /** Holt {@link XDispatchProvider} Interface von o. */
  public static XDispatchProvider XDispatchProvider(Object o)
  {
    return UnoRuntime.queryInterface(XDispatchProvider.class, o);
  }

  /** Holt {@link XDispatchHelper} Interface von o. */
  public static XDispatchHelper XDispatchHelper(Object o)
  {
    return UnoRuntime.queryInterface(XDispatchHelper.class, o);
  }

  /** Holt {@link XURLTransformer} Interface von o. */
  public static XURLTransformer XURLTransformer(Object o)
  {
    return UnoRuntime.queryInterface(XURLTransformer.class, o);
  }

  /** Holt {@link XSelectionSupplier} Interface von o. */
  public static XSelectionSupplier XSelectionSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XSelectionSupplier.class, o);
  }

  /** Holt {@link XInterface} Interface von o. */
  public static XInterface XInterface(Object o)
  {
    return UnoRuntime.queryInterface(XInterface.class, o);
  }

  /** Holt {@link XKeysSupplier} Interface von o. */
  public static XKeysSupplier XKeysSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XKeysSupplier.class, o);
  }

  /** Holt {@link XComponentLoader} Interface von o. */
  public static XComponentLoader XComponentLoader(Object o)
  {
    return UnoRuntime.queryInterface(XComponentLoader.class, o);
  }

  /** Holt {@link XBrowseNodeFactory} Interface von o. */
  public static XBrowseNodeFactory XBrowseNodeFactory(Object o)
  {
    return UnoRuntime.queryInterface(XBrowseNodeFactory.class, o);
  }

  /** Holt {@link XBookmarksSupplier} Interface von o. */
  public static XBookmarksSupplier XBookmarksSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XBookmarksSupplier.class, o);
  }

  /** Holt {@link XNameContainer} Interface von o. */
  public static XNameContainer XNameContainer(Object o)
  {
    return UnoRuntime.queryInterface(XNameContainer.class, o);
  }

  /** Holt {@link XMultiServiceFactory} Interface von o. */
  public static XMultiServiceFactory XMultiServiceFactory(Object o)
  {
    return UnoRuntime.queryInterface(XMultiServiceFactory.class, o);
  }

  /** Holt {@link XDesktop} Interface von o. */
  public static XDesktop XDesktop(Object o)
  {
    return UnoRuntime.queryInterface(XDesktop.class, o);
  }

  /** Holt {@link XChangesBatch} Interface von o. */
  public static XChangesBatch XChangesBatch(Object o)
  {
    return UnoRuntime.queryInterface(XChangesBatch.class, o);
  }

  /** Holt {@link XNameAccess} Interface von o. */
  public static XNameAccess XNameAccess(Object o)
  {
    return UnoRuntime.queryInterface(XNameAccess.class, o);
  }

  /** Holt {@link XFilePicker} Interface von o. */
  public static XFilePicker XFilePicker(Object o)
  {
    return UnoRuntime.queryInterface(XFilePicker.class, o);
  }

  /** Holt {@link XFolderPicker} Interface von o. */
  public static XFolderPicker XFolderPicker(Object o)
  {
    return UnoRuntime.queryInterface(XFolderPicker.class, o);
  }

  /** Holt {@link XToolkit} Interface von o. */
  public static XToolkit XToolkit(Object o)
  {
    return UnoRuntime.queryInterface(XToolkit.class, o);
  }

  /** Holt {@link XWindow} Interface von o. */
  public static XWindow XWindow(Object o)
  {
    return UnoRuntime.queryInterface(XWindow.class, o);
  }

  /** Holt {@link XToolkit} Interface von o. */
  public static XWindowPeer XWindowPeer(Object o)
  {
    return UnoRuntime.queryInterface(XWindowPeer.class, o);
  }

  /** Holt {@link XToolbarController} Interface von o. */
  public static XToolbarController XToolbarController(Object o)
  {
    return UnoRuntime.queryInterface(XToolbarController.class, o);
  }

  /** Holt {@link com.sun.star.text.XParagraphCursor} Interface von o. */
  public static XParagraphCursor XParagraphCursor(Object o)
  {
    return UnoRuntime.queryInterface(XParagraphCursor.class, o);
  }

  /** Holt {@link com.sun.star.lang.XServiceInf} Interface von o. */
  public static XServiceInfo XServiceInfo(Object o)
  {
    return UnoRuntime.queryInterface(XServiceInfo.class, o);
  }

  /** Holt {@link XUpdatable} Interface von o. */
  public static XUpdatable XUpdatable(Object o)
  {
    return UnoRuntime.queryInterface(XUpdatable.class, o);
  }

  /** Holt {@link XTextViewCursorSupplier} Interface von o. */
  public static XTextViewCursorSupplier XTextViewCursorSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextViewCursorSupplier.class, o);
  }

  /** Holt {@link XDrawPageSupplier} Interface von o. */
  public static XDrawPageSupplier XDrawPageSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XDrawPageSupplier.class, o);
  }

  /** Holt {@link XContentEnumerationAccess} Interface von o. */
  public static XContentEnumerationAccess XContentEnumerationAccess(Object o)
  {
    return UnoRuntime.queryInterface(XContentEnumerationAccess.class, o);
  }

  /** Holt {@link XControlShape} Interface von o. */
  public static XControlShape XControlShape(Object o)
  {
    return UnoRuntime.queryInterface(XControlShape.class, o);
  }

  /** Holt {@link XTextRange} Interface von o. */
  public static XTextRange XTextRange(Object o)
  {
    return UnoRuntime.queryInterface(XTextRange.class, o);
  }

  /** Holt {@link XDevice} Interface von o. */
  public static XDevice XDevice(Object o)
  {
    return UnoRuntime.queryInterface(XDevice.class, o);
  }

  /** Holt {@link com.sun.star.frame.XFramesSupplier} Interface von o. */
  public static XFramesSupplier XFramesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XFramesSupplier.class, o);
  }

  /** Holt {@link XDocumentPropertiesSupplier} Interface von o. */
  public static XDocumentPropertiesSupplier XDocumentPropertiesSupplier(
      Object o)
  {
    return UnoRuntime.queryInterface(XDocumentPropertiesSupplier.class, o);
  }

  /** Holt {@link com.sun.star.document.XDocumentProperties} Interface von o. */
  public static XDocumentProperties XDocumentProperties(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentProperties.class, o);
  }

  /** Holt {@link XDispatchProviderInterception} Interface von o. */
  public static XDispatchProviderInterception XDispatchProviderInterception(
      Object o)
  {
    return UnoRuntime.queryInterface(XDispatchProviderInterception.class, o);
  }

  /** Holt {@link XTextCursor} Interface von o. */
  public static XTextCursor XTextCursor(Object o)
  {
    return UnoRuntime.queryInterface(XTextCursor.class, o);
  }

  /** Holt {@link com.sun.star.text.XPageCursor} Interface von o. */
  public static XPageCursor XPageCursor(Object o)
  {
    return UnoRuntime.queryInterface(XPageCursor.class, o);
  }

  /** Holt {@link XNotifyingDispatch} Interface von o. */
  public static XNotifyingDispatch XNotifyingDispatch(Object o)
  {
    return UnoRuntime.queryInterface(XNotifyingDispatch.class, o);
  }

  /** Holt {@link com.sun.star.style.XStyleFamiliesSupplier} Interface von o. */
  public static XStyleFamiliesSupplier XStyleFamiliesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, o);
  }

  /** Holt {@link com.sun.star.style.XStyle} Interface von o. */
  public static XStyle XStyle(Object o)
  {
    return UnoRuntime.queryInterface(XStyle.class, o);
  }

  /** Holt {@link XTextRangeCompare} Interface von o. */
  public static XTextRangeCompare XTextRangeCompare(Object o)
  {
    return UnoRuntime.queryInterface(XTextRangeCompare.class, o);
  }

  /** Holt {@link com.sun.star.text.XTextSection} Interface von o. */
  public static XTextSection XTextSection(Object o)
  {
    return UnoRuntime.queryInterface(XTextSection.class, o);
  }

  /** Holt {@link com.sun.star.text.XTextSectionSupplier} Interface von o. */
  public static XTextSectionsSupplier XTextSectionsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextSectionsSupplier.class, o);
  }

  /** Holt {@link com.sun.star.ui.XAcceleratorConfiguration} Interface von o. */
  public static XAcceleratorConfiguration XAcceleratorConfiguration(Object o)
  {
    return UnoRuntime.queryInterface(XAcceleratorConfiguration.class, o);
  }

  /**
   * Holt {@link com.sun.star.ui.XUIConfigurationPersistence} Interface von o.
   */
  public static XUIConfigurationPersistence XUIConfigurationPersistence(
      Object o)
  {
    return UnoRuntime.queryInterface(XUIConfigurationPersistence.class, o);
  }

  /**
   * Holt {@link com.sun.star.ui.XModuleUIConfigurationManager} Interface von o.
   */
  public static XModuleUIConfigurationManager XModuleUIConfigurationManager(
      Object o)
  {
    return UnoRuntime.queryInterface(XModuleUIConfigurationManager.class, o);
  }

  /** Holt {@link com.sun.star.ui.XUIConfigurationManager} Interface von o. */
  public static XUIConfigurationManager XUIConfigurationManager(Object o)
  {
    return UnoRuntime.queryInterface(XUIConfigurationManager.class, o);
  }

  /** Holt {@link XIndexContainer} Interface von o. */
  public static XIndexContainer XIndexContainer(Object o)
  {
    return UnoRuntime.queryInterface(XIndexContainer.class, o);
  }

  /** Holt {@link com.sun.star.frame.XFrame} Interface von o. */
  public static XFrame XFrame(Object o)
  {
    return UnoRuntime.queryInterface(XFrame.class, o);
  }

  /** Holt {@link com.sun.star.style.XTextColumns} Interface von o. */
  public static XTextColumns XTextColumns(Object o)
  {
    return UnoRuntime.queryInterface(XTextColumns.class, o);
  }

  /** Holt {@link com.sun.star.style.XStyleLoader} Interface von o. */
  public static XStyleLoader XStyleLoader(Object o)
  {
    return UnoRuntime.queryInterface(XStyleLoader.class, o);
  }

  /** Holt {@link XPropertyState} Interface von o. */
  public static XPropertyState XPropertyState(Object o)
  {
    return UnoRuntime.queryInterface(XPropertyState.class, o);
  }

  /** Holt {@link XStringSubstitution} Interface von o. */
  public static XStringSubstitution XStringSubstitution(Object o)
  {
    return UnoRuntime.queryInterface(XStringSubstitution.class, o);
  }

  /** Holt {@link XRowSet} Interface von o. */
  public static XRowSet XRowSet(Object o)
  {
    return UnoRuntime.queryInterface(XRowSet.class, o);
  }

  /** Holt {@link XDocumentMetadataAccess} Interface von o. */
  public static XDocumentMetadataAccess XDocumentMetadataAccess(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentMetadataAccess.class, o);
  }

  /** Holt {@link XRefreshable} Interface von o. */
  public static XRefreshable XRefreshable(Object o)
  {
    return UnoRuntime.queryInterface(XRefreshable.class, o);
  }

  // ACHTUNG: Interface-Methoden fangen hier mit einem grossen X an!

  /**
   * Liefert eine vorgeparste UNO-URL von urlStr.
   * 
   * @param urlStr
   * @return vorgeparste UNO-URL von urlStr.
   * @author christoph.lutz
   */
  public static com.sun.star.util.URL getParsedUNOUrl(String urlStr)
  {
    com.sun.star.util.URL[] unoURL = new com.sun.star.util.URL[] {
        new com.sun.star.util.URL() };
    unoURL[0].Complete = urlStr;
    if (urlTransformer != null)
      urlTransformer.parseStrict(unoURL);

    return unoURL[0];
  }

  /**
   * Liefert ein Service ConfigurationAccess mit dem der lesende Zugriff auf die
   * OOo-Configuration ab dem Knoten nodepath ermöglicht wird oder null, wenn
   * der Service nicht erzeugt werden kann.
   * 
   * @param nodepath
   *          Beschreibung des Knotens des Konfigurationsbaumes, der als neue
   *          Wurzel zurückgeliefert werden soll. Ein nodepath ist z.B.
   *          "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering"
   * @return ein ConfigurationUpdateAccess mit der Wurzel an dem Knoten nodepath
   *         oder null, falls der Service nicht erzeugt werden kann (wenn z.B.
   *         der Knoten nodepath nicht existiert).
   * @author christoph.lutz
   */
  public static XNameAccess getConfigurationAccess(String nodepath)
  {
    PropertyValue[] props = new PropertyValue[] { new PropertyValue() };
    props[0].Name = "nodepath";
    props[0].Value = nodepath;
    Object confProv = getConfigurationProvider();
    try
    {
      return UNO.XNameAccess(
          UNO.XMultiServiceFactory(confProv).createInstanceWithArguments(
              "com.sun.star.configuration.ConfigurationAccess", props));
    } catch (Exception e)
    {
    }
    return null;
  }

  /**
   * Liefert ein Service ConfigurationUpdateAccess mit dem der lesende und
   * schreibende Zugriff auf die OOo-Configuration ab dem Knoten nodepath
   * ermöglicht wird oder null wenn der Service nicht erzeugt werden kann.
   * 
   * @param nodepath
   *          Beschreibung des Knotens des Konfigurationsbaumes, der als neue
   *          Wurzel zurückgeliefert werden soll. Ein nodepath ist z.B.
   *          "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering"
   * @return ein ConfigurationUpdateAccess mit der Wurzel an dem Knoten nodepath
   *         oder null, falls der Service nicht erzeugt werden kann (wenn z.B.
   *         der Knoten nodepath nicht existiert).
   * @author christoph.lutz
   */
  public static XChangesBatch getConfigurationUpdateAccess(String nodepath)
  {
    PropertyValue[] props = new PropertyValue[] { new PropertyValue() };
    props[0].Name = "nodepath";
    props[0].Value = nodepath;
    Object confProv = getConfigurationProvider();
    try
    {
      return UNO.XChangesBatch(
          UNO.XMultiServiceFactory(confProv).createInstanceWithArguments(
              "com.sun.star.configuration.ConfigurationUpdateAccess", props));
    } catch (Exception e)
    {
    }
    return null;
  }

  /**
   * Liefert den configurationProvider, mit dem der Zugriff auf die
   * Konfiguration von OOo ermöglicht wird.
   * 
   * @return ein neuer configurationProvider
   */
  private static Object getConfigurationProvider()
  {
    if (configurationProvider == null)
      configurationProvider = createUNOService(
          "com.sun.star.configuration.ConfigurationProvider");
    return configurationProvider;
  }

  /**
   * Liefert den shortcutManager zu der OOo Komponente component zurück.
   * 
   * @param component
   *          die OOo Komponente zu der der ShortcutManager geliefert werden
   *          soll z.B "com.sun.star.text.TextDocument"
   * @return der shortcutManager zur OOo Komponente component oder null falls
   *         kein shortcutManager erzeugt werden kann.
   * 
   */
  public static XAcceleratorConfiguration getShortcutManager(String component)
  {
    // XModuleUIConfigurationManagerSupplier moduleUICfgMgrSupplier
    XModuleUIConfigurationManagerSupplier moduleUICfgMgrSupplier = UNO
        .XModuleUIConfigurationManagerSupplier(UNO.createUNOService(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier"));

    if (moduleUICfgMgrSupplier == null)
      return null;

    // XCUIConfigurationManager moduleUICfgMgr
    XUIConfigurationManager moduleUICfgMgr = null;

    try
    {
      moduleUICfgMgr = moduleUICfgMgrSupplier
          .getUIConfigurationManager(component);
    } catch (NoSuchElementException e)
    {
      return null;
    }

    // XAcceleratorConfiguration xAcceleratorConfiguration
    try
    {
      Method m = moduleUICfgMgr.getClass().getMethod("getShortCutManager",
          (Class[]) null);
      return UNO
          .XAcceleratorConfiguration(m.invoke(moduleUICfgMgr, (Object[]) null));
    } catch (Exception e)
    {
      return null;
    }
  }

  /**
   * Wenn hide=true ist, so wird die Eigenschaft CharHidden für range auf true
   * gesetzt und andernfalls der Standardwert (=false) für die Property
   * CharHidden wieder hergestellt. Dadurch lässt sich der Text in range
   * unsichtbar schalten bzw. wieder sichtbar schalten. Die Repräsentation von
   * unsichtbar geschaltenen Stellen erfolgt in der Art, dass OOo für den
   * unsichtbaren Textbereich ein neuen automatisch generierten Character-Style
   * anlegt, der die Eigenschaften der bisher gesetzten Styles erbt und
   * lediglich die Eigenschaft "Sichtbarkeit" auf unsichtbar setzt. Beim
   * Aufheben einer unsichtbaren Stelle sorgt das Zurücksetzen auf den
   * Standardwert dafür, dass der vorher angelegte automatische-Style wieder
   * zurück genommen wird - so ist sichergestellt, dass das Aus- und
   * Wiedereinblenden von Textbereichen keine Änderungen der bisher gesetzten
   * Styles hervorruft.
   * 
   * @param range
   *          Der Textbereich, der aus- bzw. eingeblendet werden soll.
   * @param hide
   *          hide=true blendet aus, hide=false blendet ein.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static void hideTextRange(XTextRange range, boolean hide)
  {
    String propName = "CharHidden";
    if (hide)
    {
      UNO.setProperty(range, propName, Boolean.TRUE);
      // Workaround für update Bug
      // http://qa.openoffice.org/issues/show_bug.cgi?id=78896
      UNO.setProperty(range, propName, Boolean.FALSE);
      UNO.setProperty(range, propName, Boolean.TRUE);
    } else
    {
      // Workaround für (den anderen) update Bug
      // http://qa.openoffice.org/issues/show_bug.cgi?id=103101
      // Nur das Rücksetzen auf den Standardwert reicht nicht aus. Daher erfolgt
      // vor
      // dem Rücksetzen auf den Standardwert eine explizite Einblendung.
      UNO.setProperty(range, propName, Boolean.FALSE);
      UNO.setPropertyToDefault(range, propName);
    }
  }

	public static void forEachParagraphInRange(XTextRange range, Consumer<Object> c) {
	  XTextCursor cursor = range.getText().createTextCursorByRange(range);
	  
	  for (XEnumeration par : UnoCollection.getCollection(cursor, XEnumeration.class))
	  {
	    	c.accept(par);
	  }
	}

	public static void forEachTextPortionInRange(XTextRange range, Consumer<Object> c) {
	  XTextCursor cursor = range.getText().createTextCursorByRange(range);
	  
	  for (XEnumeration parEnum : UnoCollection.getCollection(cursor, XEnumeration.class))
	  {
	  	for (Object o : UnoCollection.getCollection(parEnum, Object.class))
	  	{
	    	c.accept(o);
	  	}
	  }
	}
	
	public static Object getFirstElementInEnumeration(Object o) throws NoSuchElementException, WrappedTargetException {
		XEnumerationAccess access = UNO.XEnumerationAccess(o);
		if (access != null) {
			return access.createEnumeration().nextElement();
		}
		
		return null;
	}
}

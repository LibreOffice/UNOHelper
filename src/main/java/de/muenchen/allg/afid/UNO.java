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
*/
package de.muenchen.allg.afid;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.function.Consumer;

import org.apache.commons.lang3.ArrayUtils;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XContainerWindowProvider;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDevice;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XExtendedToolkit;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XItemList;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMenu;
import com.sun.star.awt.XNumericField;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.awt.XProgressBar;
import com.sun.star.awt.XRadioButton;
import com.sun.star.awt.XSpinField;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XToolkit2;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XUserInputInterception;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindow2;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.awt.tab.XTabPage;
import com.sun.star.awt.tab.XTabPageContainer;
import com.sun.star.awt.tab.XTabPageContainerModel;
import com.sun.star.awt.tree.XMutableTreeNode;
import com.sun.star.awt.tree.XTreeControl;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XIntrospection;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertyState;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XContentEnumerationAccess;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.container.XSet;
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
import com.sun.star.frame.XController2;
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
import com.sun.star.io.IOException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
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
import com.sun.star.sdb.DatabaseContext;
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
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.ui.XAcceleratorConfiguration;
import com.sun.star.ui.XDeck;
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
import com.sun.star.util.XCancellable;
import com.sun.star.util.XChangesBatch;
import com.sun.star.util.XCloseBroadcaster;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XModifiable2;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XRefreshable;
import com.sun.star.util.XStringSubstitution;
import com.sun.star.util.XURLTransformer;
import com.sun.star.util.XUpdatable;
import com.sun.star.view.XPrintable;
import com.sun.star.view.XSelectionSupplier;
import com.sun.star.view.XViewSettingsSupplier;

import de.muenchen.allg.afid.Utils.FindNode;
import de.muenchen.allg.util.UnoComponent;
import de.muenchen.allg.util.UnoConfiguration;
import de.muenchen.allg.util.UnoProperty;
import de.muenchen.allg.util.UnoService;

/**
 * Helper for querying Interfaces.
 */
@SuppressWarnings("squid:S00100")
public class UNO
{

  private static final String CANT_CONNECT_TO_OFFICE = "Can't connect to Office.";

  /**
   * Main component factory.
   */
  public static XMultiComponentFactory xMCF;

  /**
   * Main service factory.
   */
  public static XMultiServiceFactory xMSF;

  /**
   * Context of {@link #xMCF}.
   */
  public static XComponentContext defaultContext;

  /**
   * The global {@link DatabaseContext}.
   */
  public static XNamingService dbContext;

  /**
   * A transformer for URLs.
   */
  public static XURLTransformer urlTransformer;

  /**
   * A dispatcher.
   */
  public static XDispatchHelper dispatchHelper;

  /**
   * The main application.
   */
  public static XDesktop desktop;

  /**
   * Component based methods use this instance. Should be controlled by the user.
   */
  public static XComponent compo;

  /**
   * The root of all scripts.
   */
  public static BrowseNode scriptRoot;

  /**
   * The provider of scripts.
   */
  public static XScriptProvider masterScriptProvider;

  /**
   * Provides access to the configuration.
   */
  private static Object configurationProvider;

  private UNO()
  {
    // hide constructor
  }

  /**
   * Initialize connection to Office.
   *
   * @param connectionString
   *          The connection parameters (eg
   *          "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager")
   * @throws UnoHelperException
   *           Can't connect to Office.
   */
  public static void init(String connectionString) throws UnoHelperException
  {
    try
    {
      XComponentContext xLocalContext = Bootstrap.createInitialComponentContext(null);
      XMultiComponentFactory xLocalFactory = xLocalContext.getServiceManager();
      XUnoUrlResolver xUrlResolver = UNO.XUnoUrlResolver(UnoComponent
          .createComponentWithContext(UnoComponent.CSS_BRIDGE_UNO_URL_RESOLVER, xLocalFactory, xLocalContext));
      init(xUrlResolver.resolve(connectionString));
    } catch (Exception e)
    {
      throw new UnoHelperException(CANT_CONNECT_TO_OFFICE, e);
    }
  }

  /**
   * Initialize connection to Office.
   *
   * @param options
   *          The connection parameters (see <a href=
   *          "https://help.libreoffice.org/4.0/Common/Starting_the_Software_With_Parameters">
   *          Starting_the_Software_With_Parameters</a>)
   * @throws UnoHelperException
   *           Can't connect to Office.
   */
  public static void init(List<String> options) throws UnoHelperException
  {
    try
    {
      init(Bootstrap.bootstrap(options.toArray(new String[options.size()])).getServiceManager());
    } catch (Exception e)
    {
      throw new UnoHelperException(CANT_CONNECT_TO_OFFICE, e);
    }
  }

  /**
   * Initialize connection to Office with default parameters.
   *
   * @throws UnoHelperException
   *           Can't connect to Office.
   */
  public static void init() throws UnoHelperException
  {
    try
    {
      init(Bootstrap.bootstrap().getServiceManager());
    } catch (Exception e)
    {
      throw new UnoHelperException(CANT_CONNECT_TO_OFFICE, e);
    }
  }

  /**
   * Initialize connection to Office.
   *
   * @param remoteServiceManager
   *          Running Office instance.
   * @throws UnoHelperException
   *           Can't connect to Office.
   */
  public static void init(Object remoteServiceManager) throws UnoHelperException
  {
    try
    {
      xMCF = UNO.XMultiComponentFactory(remoteServiceManager);
      xMSF = UNO.XMultiServiceFactory(xMCF);
      defaultContext = UNO.XComponentContext(UnoProperty.getProperty(xMCF, UnoProperty.DEFAULT_CONTEXT));

      desktop = UNO.XDesktop(UnoComponent.createComponentWithContext(UnoComponent.CSS_FRAME_DESKTOP));

      XBrowseNodeFactory masterBrowseNodeFac = UnoRuntime.queryInterface(XBrowseNodeFactory.class,
          defaultContext.getValueByName("/singletons/com.sun.star.script.browse.theBrowseNodeFactory"));
      scriptRoot = new BrowseNode(masterBrowseNodeFac.createView(BrowseNodeFactoryViewTypes.MACROORGANIZER));

      XScriptProviderFactory masterProviderFac = UnoRuntime.queryInterface(XScriptProviderFactory.class,
          defaultContext.getValueByName("/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory"));
      masterScriptProvider = masterProviderFac.createScriptProvider(defaultContext);

      dbContext = UNO.XNamingService(UnoComponent.createComponentWithContext(UnoComponent.CSS_SDB_DATABASE_CONTEXT));
      urlTransformer = UNO
          .XURLTransformer(UnoComponent.createComponentWithContext(UnoComponent.CSS_UTIL_URL_TRANSFORMER));
      dispatchHelper = UNO
          .XDispatchHelper(UnoComponent.createComponentWithContext(UnoComponent.CSS_FRAME_DISPATCH_HELPER));
    } catch (IllegalArgumentException e)
    {
      throw new UnoHelperException(CANT_CONNECT_TO_OFFICE, e);
    }
  }

  /**
   * Lädt ein Dokument und setzt im Erfolgsfall {@link #compo} auf das geöffnete
   * Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros falls true wird die Ausführung von Makros
   *                    freigeschaltet.
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, boolean allowMacros)
      throws UnoHelperException
  {
      return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Lädt ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros falls true wird die Ausführung von Makros
   *                    freigeschaltet.
   * @param hidden      falls true wird das Dokument unsichtbar geöffnet
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, boolean allowMacros, boolean hidden)
      throws UnoHelperException
  {
    short allowMacrosShort = MacroExecMode.NEVER_EXECUTE;
    if (allowMacros)
    {
      allowMacrosShort = MacroExecMode.ALWAYS_EXECUTE_NO_WARN;
    }
    return loadComponentFromURL(url, asTemplate, allowMacrosShort, hidden);
  }

  /**
   * Lädt ein Dokument und setzt im Erfolgsfall {@link #compo} auf das geöffnete
   * Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros eine der Konstanten aus {@link MacroExecMode}.
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, short allowMacros)
      throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Lädt ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros eine der Konstanten aus {@link MacroExecMode}.
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
    short allowMacros, boolean hidden)
      throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, hidden,
        new PropertyValue[] {});
  }

  /**
   * Lädt ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros falls true wird die Ausführung von Makros
   *                    freigeschaltet.
   * @param args        zusätzliche Parameter für
   *                    XComponentLoader.loadComponentFromUrl
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
    boolean allowMacros, PropertyValue... args)
      throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate,
      (allowMacros) ? MacroExecMode.ALWAYS_EXECUTE_NO_WARN
        : MacroExecMode.NEVER_EXECUTE,
      false, args);
  }

  /**
   * Lädt ein Dokument abhängig von hidden sichtbar oder unsichtbar und setzt im
   * Erfolgsfall {@link #compo} auf das geöffnete Dokument.
   * 
   * @param url         die URL des zu ladenden Dokuments, z.B.
   *                    "file:///C:/temp/footest.odt" oder
   *                    "private:factory/swriter" (für ein leeres).
   * @param asTemplate  falls true wird das Dokument als Vorlage behandelt und
   *                    ein neues unbenanntes Dokument erzeugt.
   * @param allowMacros eine der Konstanten aus {@link MacroExecMode}.
   * @param args        zusätzliche Parameter für
   *                    XComponentLoader.loadComponentFromUrl
   * @return das geöffnete Dokument
   * @throws UnoHelperException
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate,
    short allowMacros, boolean hidden, PropertyValue... args)
      throws UnoHelperException
  {
    try
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

      arguments = ArrayUtils.addAll(arguments, args);

      XComponent lc = loader.loadComponentFromURL(url, "_blank",
          FrameSearchFlag.CREATE, arguments);
      if (lc != null)
        compo = lc;
      return lc;
    }
    catch (IllegalArgumentException | IOException e)
    {
      throw new UnoHelperException(
          "Dokument konnte nicht geladen werden.", e);
    }
  }
  
  /**
   * Konvertiert Dateipfad in eine systemspezifische URL.
   * 
   * @param filePath
   * @return eine für LO valide URL oder ein leerer String falls der Dateipfad nicht durch XFileIdentifierConverter konvertiert werden konnte.
   * @throws UnoHelperException
   */
  public static String convertFilePathToURL(String filePath) throws UnoHelperException {
    String resultURL = null;
    
    try
    {
      Object fileContentProvider = UNO.xMCF.createInstanceWithContext("com.sun.star.ucb.FileContentProvider",
          UNO.defaultContext);
      XFileIdentifierConverter xFileConverter = UnoRuntime.queryInterface(XFileIdentifierConverter.class,
          fileContentProvider);
      resultURL = xFileConverter.getFileURLFromSystemPath("", filePath);
    } catch (Exception e)
    {
      throw new UnoHelperException("", e);
    } 
    
    return resultURL;
  }

  /**
   * Dispatcht url auf dem aktuellen Controll von doc.
   */
  public static void dispatch(XModel doc, String url)
  {
    XDispatchProvider prov = UNO.XDispatchProvider(doc.getCurrentController().getFrame());
    dispatchHelper.executeDispatch(prov, url, "", FrameSearchFlag.SELF, new PropertyValue[]
    {});
  }

  public static XTextDocument getCurrentTextDocument()
  {
    XComponent xComponent = desktop.getCurrentComponent();

    return UnoRuntime.queryInterface(com.sun.star.text.XTextDocument.class, xComponent);
  }

  /**
   * Dispatcht url auf dem aktuellen Controll von doc mittels
   * {@link XNotifyingDispatch}, wartet auf die Benachrichtigung, dass der
   * Dispatch vollständig abgearbeitet ist, und liefert das
   * {@link DispatchResultEvent} dieser Benachrichtigung zurück oder null, wenn zu
   * url kein XNotifyingDispatch definiert ist oder der Dispatch mit disposing
   * abgebrochen wurde.
   */
  public static DispatchResultEvent dispatchAndWait(XTextDocument doc, String url)
  {
    if (doc == null)
      return null;

    URL unoUrl = getParsedUNOUrl(url);

    XDispatchProvider prov = UNO.XDispatchProvider(doc.getCurrentController().getFrame());
    if (prov == null)
      return null;

    XNotifyingDispatch disp = UNO.XNotifyingDispatch(prov.queryDispatch(unoUrl, "", FrameSearchFlag.SELF));
    if (disp == null)
      return null;

    final boolean[] lock = new boolean[]
    { true };
    final DispatchResultEvent[] resultEvent = new DispatchResultEvent[]
    { null };

    disp.dispatchWithNotification(unoUrl, new PropertyValue[]
    {}, new XDispatchResultListener()
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
   *                                   ist ein Objekt, das entweder
   *                                   {@link XScriptProvider} oder
   *                                   {@link XScriptProviderSupplier}
   *                                   implementiert. Dies kann z.B. ein
   *                                   TextDocument sein. Soll einfach nur ein
   *                                   Skript aus dem gesamten Skript-Baum
   *                                   ausgeführt werden, kann die Funktion
   *                                   {@link #executeGlobalMacro(String, Object[])}
   *                                   verwendet werden, die diesen Parameter
   *                                   nicht erfordert. ACHTUNG! Es wird nicht
   *                                   zwangsweise der übergebene
   *                                   scriptProviderOrSupplier verwendet um das
   *                                   Skript auszuführen. Er stellt nur den
   *                                   Einstieg in den Skript-Baum dar.
   * @param macroName
   *                                   ist der Name des Makros. Der Name kann
   *                                   optional durch "." abgetrennte Bezeichner
   *                                   für Bibliotheken/Module vorangestellt
   *                                   haben. Es sind also sowohl "Foo" als auch
   *                                   "Module1.Foo" und "Standard.Module1.Foo"
   *                                   erlaubt. Wenn kein passendes Makro gefunden
   *                                   wird, wird zuerst versucht,
   *                                   case-insensitive danach zu suchen. Falls
   *                                   dabei ebenfalls kein Makro gefunden wird,
   *                                   wird eine {@link RuntimeException}
   *                                   geworfen.
   * @param args
   *                                   die Argumente, die dem Makro übergeben
   *                                   werden sollen.
   * @param location
   *                                   eine Liste aller erlaubten locations
   *                                   ("application", "share", "document") für
   *                                   das Makro. Bei der Suche wird zuerst ein
   *                                   case-sensitive Match in allen gelisteten
   *                                   locations gesucht, bevor die
   *                                   case-insensitive Suche versucht wird. Durch
   *                                   Verwendung der exakten
   *                                   Gross-/Kleinschreibung des Makros und
   *                                   korrekte Ordnung der location Liste lässt
   *                                   sich also immer das richtige Makro
   *                                   selektieren.
   * @throws RuntimeException
   *                            wenn entweder kein passendes Makro gefunden wurde,
   *                            oder scriptProviderOrSupplier weder
   *                            {@link XScriptProvider} noch
   *                            {@link XScriptProviderSupplier} implementiert.
   * @return den Rückgabewert des Makros.
   * @throws UnoHelperException 
   */
  public static Object executeMacro(Object scriptProviderOrSupplier,
      String macroName, Object[] args, String[] location)
      throws UnoHelperException
  {
    XScriptProvider provider = UnoRuntime.queryInterface(XScriptProvider.class, scriptProviderOrSupplier);
    if (provider == null)
    {
      XScriptProviderSupplier supp = UnoRuntime.queryInterface(XScriptProviderSupplier.class, scriptProviderOrSupplier);
      if (supp == null)
        throw new RuntimeException("Übergebenes Objekt ist weder XScriptProvider noch XScriptProviderSupplier");
      provider = supp.getScriptProvider();
    }

    XBrowseNode root = UnoRuntime.queryInterface(XBrowseNode.class, provider);
    /*
     * Wir übergeben NICHT provider als drittes Argument, sondern lassen
     * Internal.executeMacroInternal den provider selbst bestimmen. Das hat keinen
     * besonderen Grund. Es erscheint einfach nur etwas robuster, den
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
   *                    ist der Name des Makros. Der Name kann optional durch "."
   *                    abgetrennte Bezeichner für Bibliotheken/Module
   *                    vorangestellt haben. Es sind also sowohl "Foo" als auch
   *                    "Module1.Foo" und "Standard.Module1.Foo" erlaubt. Wenn
   *                    kein passendes Makro gefunden wird, wird zuerst versucht,
   *                    case-insensitive danach zu suchen. Falls dabei ebenfalls
   *                    kein Makro gefunden wird, wird eine
   *                    {@link RuntimeException} geworfen.
   * @param args
   *                    die Argumente, die dem Makro übergeben werden sollen.
   * @throws RuntimeException
   *                            wenn kein passendes Makro gefunden wurde.
   * @return den Rückgabewert des Makros.
   * @throws UnoHelperException 
   */
  public static Object executeGlobalMacro(String macroName, Object[] args)
      throws UnoHelperException
  {
    final String[] userAndShare = new String[]
    { "application", "share" };
    return Utils.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), userAndShare);
  }

  /**
   * Ruft ein Makro aus dem gesamten Makro-Baum auf.
   * 
   * @param macroName
   *                    ist der Name des Makros. Der Name kann optional durch "."
   *                    abgetrennte Bezeichner für Bibliotheken/Module
   *                    vorangestellt haben. Es sind also sowohl "Foo" als auch
   *                    "Module1.Foo" und "Standard.Module1.Foo" erlaubt. Wenn
   *                    kein passendes Makro gefunden wird, wird zuerst versucht,
   *                    case-insensitive danach zu suchen. Falls dabei ebenfalls
   *                    kein Makro gefunden wird, wird eine
   *                    {@link RuntimeException} geworfen.
   * @param args
   *                    die Argumente, die dem Makro übergeben werden sollen.
   * @param location
   *                    eine Liste aller erlaubten locations ("application",
   *                    "share", "document") für das Makro. Bei der Suche wird
   *                    zuerst ein case-sensitive Match in allen gelisteten
   *                    locations gesucht, bevor die case-insenstive Suche
   *                    versucht wird. Durch Verwendung der exakten
   *                    Gross-/Kleinschreibung des Makros und korrekte Ordnung der
   *                    location Liste lässt sich also immer das richtige Makro
   *                    selektieren.
   * @throws RuntimeException
   *                            wenn kein passendes Makro gefunden wurde.
   * @return den Rückgabewert des Makros.
   * @throws UnoHelperException 
   */
  public static Object executeMacro(String macroName, Object[] args,
      String[] location) throws UnoHelperException
  {
    return Utils.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), location);
  }

  /**
   * Remove the part after last occurrence of ' -'.
   *
   * @param str
   *          The string which may contain ' -'.
   * @return A string without the last part.
   */
  public static String stripOpenOfficeFromWindowName(String str)
  {
    int idx = str.lastIndexOf(" -");
    if (idx > 0)
    {
      str = str.substring(0, idx);
    }
    return str;
  }

  /**
   * Get the version of LibreOffice.
   *
   * @return The concatenation of the values {@link UnoProperty#OO_SETUP_VERSION_ABOUT_BOX} and
   *         {@link UnoProperty#OO_SETUP_EXTENSION or null if configuration can't be accessed.
   */
  public static String getOOoVersion()
  {
    try
    {
      return ""
          + UnoConfiguration.getConfiguration("/org.openoffice.Setup/Product", UnoProperty.OO_SETUP_VERSION_ABOUT_BOX)
          + UnoConfiguration.getConfiguration("/org.openoffice.Setup/Product", UnoProperty.OO_SETUP_EXTENSION);
    } catch (UnoHelperException e)
    {
      return null;
    }
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT, dessen Name
   * nameToFind ist (kann durch "." abgetrennte Pfadangabe im Skript-Baum enthalten). Siehe
   * {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String[])}
   * .
   * 
   * @return den gefundenen Knoten oder null falls keiner gefunden.
   * @throws UnoHelperException
   */
  public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode, String prefix, String nameToFind,
      boolean caseSensitive) throws UnoHelperException
  {
    XBrowseNodeAndXScriptProvider x = findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind,
        caseSensitive);
    return x.getXBrowseNode();
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT,
   * dessen Name nameToFind ist (kann durch "." abgetrennte Pfadangabe im
   * Skript-Baum enthalten). Siehe
   * {@link #findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode, String, String, boolean, String[])}
   * .
   * 
   * @return den gefundenen Knoten, sowie den nächsten Vorfahren, der
   *         XScriptProvider implementiert (oder den Knoten selbst, falls dieser
   *         XScriptProvider implementiert). Falls kein entsprechender Knoten oder
   *         Vorfahre gefunden wurde, wird der entsprechende Wert als null
   *         geliefert.
   * @throws UnoHelperException 
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode,
      String prefix, String nameToFind, boolean caseSensitive)
      throws UnoHelperException
  {
    return findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind, caseSensitive, null);
  }

  /**
   * Durchsucht einen {@link XBrowseNode} Baum nach einem Blatt vom Typ SCRIPT,
   * dessen Name nameToFind ist (kann durch "." abgetrennte Pfadangabe im
   * Skript-Baum enthalten).
   * .
   * 
   * @param xBrowseNode
   *                        Wurzel des zu durchsuchenden Baums.
   * @param prefix
   *                        wird dem Namen jedes Knoten vorangestellt. Dies wird
   *                        verwendet, wenn xBrowseNode nicht die Wurzel ist.
   * @param nameToFind
   *                        der zu suchende Name.
   * @param caseSensitive
   *                        falls true, so wird Gross-/Kleinschreibung
   *                        berücksichtigt bei der Suche.
   * @param location
   *                        Es gelten nur Knoten als Treffer, die ein "URI"
   *                        Property haben, das eine location enthält die einem
   *                        String in der <code>location</code> Liste entspricht.
   *                        Mögliche locations sind "document", "application" und
   *                        "share". Falls <code>location==null</code>, so wird
   *                        {"document", "application", "share"} angenommen.
   * @return den gefundenen Knoten, sowie den nächsten Vorfahren, der
   *         XScriptProvider implementiert. Falls kein entsprechender Knoten oder
   *         Vorfahre gefunden wurde, wird der entsprechende Wert als null
   *         geliefert.
   * @throws UnoHelperException 
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode,
      String prefix, String nameToFind, boolean caseSensitive,
      String[] location) throws UnoHelperException
  {
    String[] loc;
    if (location == null)
    {
      loc = new String[]
      { "document", "application", "share" };
    } else
    {
      loc = location;
    }
    List<FindNode> found = new LinkedList<>();
    List<String> prefixVec = new ArrayList<>();
    List<String> prefixLCVec = new ArrayList<>();
    String[] prefixArr = prefix.split("\\.");
    for (int i = 0; i < prefixArr.length; ++i)
    {
      if (!prefixArr[i].isEmpty())
      {
        prefixVec.add(prefixArr[i]);
        prefixLCVec.add(prefixArr[i].toLowerCase());
      }
    }

    String[] nameToFindArrPre = nameToFind.split("\\.");
    int i1 = 0;
    while (i1 < nameToFindArrPre.length && nameToFindArrPre[i1].isEmpty())
    {
      ++i1;
    }
    int i2 = nameToFindArrPre.length - 1;
    while (i2 >= i1 && nameToFindArrPre[i2].isEmpty())
    {
      --i2;
    }
    ++i2;
    String[] nameToFindArr = new String[i2 - i1];
    String[] nameToFindLCArr = new String[i2 - i1];
    for (int i = 0; i < nameToFindArr.length; ++i)
    {
      nameToFindArr[i] = nameToFindArrPre[i1];
      nameToFindLCArr[i] = nameToFindArrPre[i1].toLowerCase();
      ++i1;
    }

    Utils.findBrowseNodeTreeLeavesAndScriptProviders(new BrowseNode(xBrowseNode), prefixVec, prefixLCVec, nameToFindArr,
        nameToFindLCArr, loc, null, found);

    if (found.isEmpty())
    {
      return new XBrowseNodeAndXScriptProvider(null, null);
    }

    Utils.FindNode findNode = found.get(0);

    if (caseSensitive && !findNode.isCaseCorrect())
    {
      return new XBrowseNodeAndXScriptProvider(null, null);
    }

    return new XBrowseNodeAndXScriptProvider(findNode.getXBrowseNode(), findNode.getXScriptProvider());
  }

  /**
   * @see UnoProperty#getProperty(Object, String)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static Object getProperty(Object o, String propName)
      throws UnoHelperException
  {
    return UnoProperty.getProperty(o, propName);
  }

  /**
   * @see UnoProperty#getPropertyByPropertyValues(PropertyValue[], String)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static String getPropertyByPropertyValues(PropertyValue[] propertyValues,
      String propertyName)
  {
    return UnoProperty.getPropertyByPropertyValues(propertyValues, propertyName);
  }

  /**
   * @see UnoProperty#setProperty(Object, String, Object)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static Object setProperty(Object o, String propName, Object propVal)
      throws UnoHelperException
  {
    return UnoProperty.setProperty(o, propName, propVal);
  }

  /**
   * @see UnoProperty#setPropertyToDefault(Object, String)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static Object setPropertyToDefault(Object o, String propName)
  {
    try
    {
      return UnoProperty.setPropertyToDefault(o, propName);
    } catch (UnoHelperException e)
    {
      return null;
    }
  }

  /**
   * @see UnoService#supportsService(Object, String)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static boolean supportsService(Object service, String serviceName)
  {
    return UnoService.supportsService(service, serviceName);
  }

  /**
   * @see UnoComponent#createComponentWithContext(String)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public static Object createUNOService(String serviceName)
  {
    return UnoComponent.createComponentWithContext(serviceName);
  }

  /**
   * Get {@link XComponentContext} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XComponentContext Interface.
   */
  public static XComponentContext XComponentContext(Object o)
  {
    return UnoRuntime.queryInterface(XComponentContext.class, o);
  }

  /**
   * Get {@link XSingleServiceFactory} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XSingleServiceFactory Interface.
   */
  public static XSingleServiceFactory XSingleServiceFactory(Object o)
  {
    return UnoRuntime.queryInterface(XSingleServiceFactory.class, o);
  }

  /**
   * Get {@link XViewSettingsSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XViewSettingsSupplier Interface.
   */
  public static XViewSettingsSupplier XViewSettingsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XViewSettingsSupplier.class, o);
  }

  /**
   * Get {@link XStorable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStorable Interface.
   */
  public static XStorable XStorable(Object o)
  {
    return UnoRuntime.queryInterface(XStorable.class, o);
  }

  /**
   * Get {@link XAutoTextContainer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XAutoTextContainer Interface.
   */
  public static XAutoTextContainer XAutoTextContainer(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextContainer.class, o);
  }

  /**
   * Get {@link XAutoTextGroup} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XAutoTextGroup Interface.
   */
  public static XAutoTextGroup XAutoTextGroup(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextGroup.class, o);
  }

  /**
   * Get {@link XAutoTextEntry} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XAutoTextEntry Interface.
   */
  public static XAutoTextEntry XAutoTextEntry(Object o)
  {
    return UnoRuntime.queryInterface(XAutoTextEntry.class, o);
  }

  /**
   * Get {@link XShape} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XShape Interface.
   */
  public static XShape XShape(Object o)
  {
    return UnoRuntime.queryInterface(XShape.class, o);
  }

  /**
   * Get {@link XTopWindow} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTopWindow Interface.
   */
  public static XTopWindow XTopWindow(Object o)
  {
    return UnoRuntime.queryInterface(XTopWindow.class, o);
  }

  /**
   * Get {@link XCloseable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XCloseable Interface.
   */
  public static XCloseable XCloseable(Object o)
  {
    return UnoRuntime.queryInterface(XCloseable.class, o);
  }

  /**
   * Get {@link XCloseBroadcaster} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XCloseBroadcaster Interface.
   */
  public static XCloseBroadcaster XCloseBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XCloseBroadcaster.class, o);
  }

  /**
   * Get {@link XStorageBasedDocument} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStorageBasedDocument Interface.
   */
  public static XStorageBasedDocument XStorageBasedDocument(Object o)
  {
    return UnoRuntime.queryInterface(XStorageBasedDocument.class, o);
  }

  /**
   * Get {@link XWindow2} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XWindow2 Interface.
   */
  public static XWindow2 XWindow2(Object o)
  {
    return UnoRuntime.queryInterface(XWindow2.class, o);
  }

  /**
   * Get {@link XTextFramesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextFramesSupplier Interface.
   */
  public static XTextFramesSupplier XTextFramesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextFramesSupplier.class, o);
  }

  /**
   * Get {@link XTextGraphicObjectsSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextGraphicObjectsSupplier Interface.
   */
  public static XTextGraphicObjectsSupplier XTextGraphicObjectsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, o);
  }

  /**
   * Get {@link XTextFrame} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextFrame Interface.
   */
  public static XTextFrame XTextFrame(Object o)
  {
    return UnoRuntime.queryInterface(XTextFrame.class, o);
  }

  /**
   * Get {@link XTextFieldsSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextFieldsSupplier Interface.
   */
  public static XTextFieldsSupplier XTextFieldsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextFieldsSupplier.class, o);
  }

  /**
   * Get {@link XComponent} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XComponent Interface.
   */
  public static XComponent XComponent(Object o)
  {
    return UnoRuntime.queryInterface(XComponent.class, o);
  }

  /**
   * Get {@link XEventBroadcaster} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XEventBroadcaster Interface.
   */
  public static XEventBroadcaster XEventBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XEventBroadcaster.class, o);
  }

  /**
   * Get {@link XBrowseNode} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XBrowseNode Interface.
   */
  public static XBrowseNode XBrowseNode(Object o)
  {
    return UnoRuntime.queryInterface(XBrowseNode.class, o);
  }

  /**
   * Get {@link XCellRangeData} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XCellRangeData Interface.
   */
  public static XCellRangeData XCellRangeData(Object o)
  {
    return UnoRuntime.queryInterface(XCellRangeData.class, o);
  }

  /**
   * Get {@link XCell} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XCell Interface.
   */
  public static XCell XCell(Object o)
  {
    return UnoRuntime.queryInterface(XCell.class, o);
  }

  /**
   * Get {@link XColumnsSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XColumnsSupplier Interface.
   */
  public static XColumnsSupplier XColumnsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XColumnsSupplier.class, o);
  }

  /**
   * Get {@link XTableColumns} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTableColumns Interface.
   */
  public static XTableColumns XTableColumns(Object o)
  {
    return UnoRuntime.queryInterface(XTableColumns.class, o);
  }

  /**
   * Get {@link XColumnRowRange} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XColumnRowRange Interface.
   */
  public static XColumnRowRange XColumnRowRange(Object o)
  {
    return UnoRuntime.queryInterface(XColumnRowRange.class, o);
  }

  /**
   * Get {@link XDeck} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDeck Interface.
   */
  public static XDeck XDeck(Object o)
  {
    return UnoRuntime.queryInterface(XDeck.class, o);
  }

  /**
   * Get {@link XIndexAccess} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XIndexAccess} Interface.
   */
  public static XIndexAccess XIndexAccess(Object o)
  {
    return UnoRuntime.queryInterface(XIndexAccess.class, o);
  }

  /**
   * Get {@link XControlContainer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XControlContainer} Interface.
   */
  public static XControlContainer XControlContainer(Object o)
  {
    return UnoRuntime.queryInterface(XControlContainer.class, o);
  }

  /**
   * Get {@link XControlModel} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XControlModel} Interface.
   */
  public static XControlModel XControlModel(Object o)
  {
    return UnoRuntime.queryInterface(XControlModel.class, o);
  }

  /**
   * Get {@link XCellRange} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XCellRange} Interfaae.
   */
  public static XCellRange XCellRange(Object o)
  {
    return UnoRuntime.queryInterface(XCellRange.class, o);
  }

  /**
   * Get {@link XUserInputInterception} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XUserInputInterception} Interfaae.
   */
  public static XUserInputInterception XUserInputInterception(Object o)
  {
    return UnoRuntime.queryInterface(XUserInputInterception.class, o);
  }

  /**
   * Get {@link XModifiable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XModifiable} Interfaae.
   */
  public static XModifiable XModifiable(Object o)
  {
    return UnoRuntime.queryInterface(XModifiable.class, o);
  }

  /**
   * Get {@link XRow} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XRow} Interfaae.
   */
  public static XRow XRow(Object o)
  {
    return UnoRuntime.queryInterface(XRow.class, o);
  }

  /**
   * Get {@link XCellRangesQuery} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XCellRangesQuery} Interfaae.
   */
  public static XCellRangesQuery XCellRangesQuery(Object o)
  {
    return UnoRuntime.queryInterface(XCellRangesQuery.class, o);
  }

  /**
   * Get {@link XDocumentDataSource} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XDocumentDataSource} Interfaae.
   */
  public static XDocumentDataSource XDocumentDataSource(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentDataSource.class, o);
  }

  /**
   * Get {@link XDataSource} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XDataSource} Interfaae.
   */
  public static XDataSource XDataSource(Object o)
  {
    return UnoRuntime.queryInterface(XDataSource.class, o);
  }

  /**
   * Get {@link XTablesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTablesSupplier} Interfaae.
   */
  public static XTablesSupplier XTablesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTablesSupplier.class, o);
  }

  /**
   * Get {@link XQueriesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XQueriesSupplier} Interfaae.
   */
  public static XQueriesSupplier XQueriesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XQueriesSupplier.class, o);
  }

  /**
   * Get {@link XNamingService} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XNamingService} Interfaae.
   */
  public static XNamingService XNamingService(Object o)
  {
    return UnoRuntime.queryInterface(XNamingService.class, o);
  }

  /**
   * Get {@link XNamed} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XNamed} Interfaae.
   */
  public static XNamed XNamed(Object o)
  {
    return UnoRuntime.queryInterface(XNamed.class, o);
  }

  /**
   * Get {@link XUnoUrlResolver} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XUnoUrlResolver} Interfaae.
   */
  public static XUnoUrlResolver XUnoUrlResolver(Object o)
  {
    return UnoRuntime.queryInterface(XUnoUrlResolver.class, o);
  }

  /**
   * Get {@link XMultiPropertySet} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XMultiPropertySet} Interfaae.
   */
  public static XMultiPropertySet XMultiPropertySet(Object o)
  {
    return UnoRuntime.queryInterface(XMultiPropertySet.class, o);
  }

  /**
   * Get {@link XPropertySet} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XPropertySet} Interfaae.
   */
  public static XPropertySet XPropertySet(Object o)
  {
    return UnoRuntime.queryInterface(XPropertySet.class, o);
  }

  /**
   * Get {@link XModel} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XModel} Interfaae.
   */
  public static XModel XModel(Object o)
  {
    return UnoRuntime.queryInterface(XModel.class, o);
  }

  /**
   * Get {@link XController} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XController} Interfaae.
   */
  public static XController XController(Object o)
  {
    return UnoRuntime.queryInterface(XController.class, o);
  }

  /**
   * Get {@link XController2} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XController2} Interfaae.
   */
  public static XController2 XController2(Object o)
  {
    return UnoRuntime.queryInterface(XController2.class, o);
  }

  /**
   * Get {@link XTextField} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTextField} Interface.
   */
  public static XTextField XTextField(Object o)
  {
    return UnoRuntime.queryInterface(XTextField.class, o);
  }

  /**
   * Get {@link XDocumentInsertable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XDocumentInsertable} Interface.
   */
  public static XDocumentInsertable XDocumentInsertable(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentInsertable.class, o);
  }

  /**
   * Get {@link XModifyBroadcaster} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XModifyBroadcaster} Interface.
   */
  public static XModifyBroadcaster XModifyBroadcaster(Object o)
  {
    return UnoRuntime.queryInterface(XModifyBroadcaster.class, o);
  }

  /**
   * Get {@link XScriptProvider} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XScriptProvider} Interface.
   */
  public static XScriptProvider XScriptProvider(Object o)
  {
    return UnoRuntime.queryInterface(XScriptProvider.class, o);
  }

  /**
   * Get {@link XSpreadsheet} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XSpreadsheet} Interface.
   */
  public static XSpreadsheet XSpreadsheet(Object o)
  {
    return UnoRuntime.queryInterface(XSpreadsheet.class, o);
  }

  /**
   * Get {@link XModuleUIConfigurationManagerSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XModuleUIConfigurationManagerSupplier} Interface.
   */
  public static XModuleUIConfigurationManagerSupplier XModuleUIConfigurationManagerSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XModuleUIConfigurationManagerSupplier.class, o);
  }

  /**
   * Get {@link XText} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XText} Interface.
   */
  public static XText XText(Object o)
  {
    return UnoRuntime.queryInterface(XText.class, o);
  }

  /**
   * Get {@link XTextTable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTextTable} Interface.
   */
  public static XTextTable XTextTable(Object o)
  {
    return UnoRuntime.queryInterface(XTextTable.class, o);
  }

  /**
   * Get {@link XDependentTextField} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XDependentTextField} Interface.
   */
  public static XDependentTextField XDependentTextField(Object o)
  {
    return UnoRuntime.queryInterface(XDependentTextField.class, o);
  }

  /**
   * Get {@link XLayoutManager} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XLayoutManager} Interface.
   */
  public static XLayoutManager XLayoutManager(Object o)
  {
    return UnoRuntime.queryInterface(XLayoutManager.class, o);
  }

  /**
   * Get {@link XColumnLocate} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XColumnLocate} Interface.
   */
  public static XColumnLocate XColumnLocate(Object o)
  {
    return UnoRuntime.queryInterface(XColumnLocate.class, o);
  }

  /**
   * Get {@link XEnumerationAccess} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XEnumerationAccess} Interface.
   */
  public static XEnumerationAccess XEnumerationAccess(Object o)
  {
    return UnoRuntime.queryInterface(XEnumerationAccess.class, o);
  }

  /**
   * Get {@link XSpreadsheetDocument} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XSpreadsheetDocument} Interface.
   */
  public static XSpreadsheetDocument XSpreadsheetDocument(Object o)
  {
    return UnoRuntime.queryInterface(XSpreadsheetDocument.class, o);
  }

  /**
   * Get {@link XPrintable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XPrintable} Interface.
   */
  public static XPrintable XPrintable(Object o)
  {
    return UnoRuntime.queryInterface(XPrintable.class, o);
  }

  /**
   * Get {@link XTabPage} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTabPage} Interface.
   */
  public static XTabPage XTabPage(Object o)
  {
    return UnoRuntime.queryInterface(XTabPage.class, o);
  }

  /**
   * Get {@link XTabPageContainer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTabPageContainer} Interface.
   */
  public static XTabPageContainer XTabPageContainer(Object o)
  {
    return UnoRuntime.queryInterface(XTabPageContainer.class, o);
  }

  /**
   * Get {@link XTabPageContainerModel} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTabPageContainerModel} Interface.
   */
  public static XTabPageContainerModel XTabPageContainerModel(Object o)
  {
    return UnoRuntime.queryInterface(XTabPageContainerModel.class, o);
  }

  /**
   * Get {@link XTextDocument} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextDocument Interface.
   */
  public static XTextDocument XTextDocument(Object o)
  {
    return UnoRuntime.queryInterface(XTextDocument.class, o);
  }

  /**
   * Get {@link XTextContent} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextContent Interface.
   */
  public static XTextContent XTextContent(Object o)
  {
    return UnoRuntime.queryInterface(XTextContent.class, o);
  }

  /**
   * Get {@link XDispatchProvider} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDispatchProvider Interface.
   */
  public static XDispatchProvider XDispatchProvider(Object o)
  {
    return UnoRuntime.queryInterface(XDispatchProvider.class, o);
  }

  /**
   * Get {@link XDispatchHelper} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDispatchHelper Interface.
   */
  public static XDispatchHelper XDispatchHelper(Object o)
  {
    return UnoRuntime.queryInterface(XDispatchHelper.class, o);
  }

  /**
   * Get {@link XURLTransformer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XURLTransformer Interface.
   */
  public static XURLTransformer XURLTransformer(Object o)
  {
    return UnoRuntime.queryInterface(XURLTransformer.class, o);
  }

  /**
   * Get {@link XSelectionSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XSelectionSupplier Interface.
   */
  public static XSelectionSupplier XSelectionSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XSelectionSupplier.class, o);
  }

  /**
   * Get {@link XInterface} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XInterface Interface.
   */
  public static XInterface XInterface(Object o)
  {
    return UnoRuntime.queryInterface(XInterface.class, o);
  }

  /**
   * Get {@link XKeysSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XKeysSupplier Interface.
   */
  public static XKeysSupplier XKeysSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XKeysSupplier.class, o);
  }

  /**
   * Get {@link XComponentLoader} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XComponentLoader Interface.
   */
  public static XComponentLoader XComponentLoader(Object o)
  {
    return UnoRuntime.queryInterface(XComponentLoader.class, o);
  }

  /**
   * Get {@link XBrowseNodeFactory} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XBrowseNodeFactory Interface.
   */
  public static XBrowseNodeFactory XBrowseNodeFactory(Object o)
  {
    return UnoRuntime.queryInterface(XBrowseNodeFactory.class, o);
  }

  /**
   * Get {@link XBookmarksSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XBookmarksSupplier Interface.
   */
  public static XBookmarksSupplier XBookmarksSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XBookmarksSupplier.class, o);
  }

  /**
   * Get {@link XNameContainer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XNameContainer Interface.
   */
  public static XNameContainer XNameContainer(Object o)
  {
    return UnoRuntime.queryInterface(XNameContainer.class, o);
  }

  /**
   * Get {@link XMultiComponentFactory} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XMultiComponentFactory Interface.
   */
  public static XMultiComponentFactory XMultiComponentFactory(Object o)
  {
    return UnoRuntime.queryInterface(XMultiComponentFactory.class, o);
  }

  /**
   * Get {@link XMultiServiceFactory} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XMultiServiceFactory Interface.
   */
  public static XMultiServiceFactory XMultiServiceFactory(Object o)
  {
    return UnoRuntime.queryInterface(XMultiServiceFactory.class, o);
  }

  /**
   * Get {@link XDesktop} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDesktop Interface.
   */
  public static XDesktop XDesktop(Object o)
  {
    return UnoRuntime.queryInterface(XDesktop.class, o);
  }

  /**
   * Get {@link XChangesBatch} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XChangesBatch Interface.
   */
  public static XChangesBatch XChangesBatch(Object o)
  {
    return UnoRuntime.queryInterface(XChangesBatch.class, o);
  }

  /**
   * Get {@link XNameAccess} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XNameAccess Interface.
   */
  public static XNameAccess XNameAccess(Object o)
  {
    return UnoRuntime.queryInterface(XNameAccess.class, o);
  }

  /**
   * Get {@link XFilePicker} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XFilePicker Interface.
   */
  public static XFilePicker XFilePicker(Object o)
  {
    return UnoRuntime.queryInterface(XFilePicker.class, o);
  }

  /**
   * Get {@link XFolderPicker} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XFolderPicker Interface.
   */
  public static XFolderPicker XFolderPicker(Object o)
  {
    return UnoRuntime.queryInterface(XFolderPicker.class, o);
  }

  /**
   * Get {@link XToolkit} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XToolkit Interface.
   */
  public static XToolkit XToolkit(Object o)
  {
    return UnoRuntime.queryInterface(XToolkit.class, o);
  }

  /**
   * Get {@link XWindow} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XWindow Interface.
   */
  public static XWindow XWindow(Object o)
  {
    return UnoRuntime.queryInterface(XWindow.class, o);
  }

  /**
   * Get {@link XWindowPeer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XWindowPeer Interface.
   */
  public static XWindowPeer XWindowPeer(Object o)
  {
    return UnoRuntime.queryInterface(XWindowPeer.class, o);
  }

  /**
   * Get {@link XToolbarController} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XToolbarController Interface.
   */
  public static XToolbarController XToolbarController(Object o)
  {
    return UnoRuntime.queryInterface(XToolbarController.class, o);
  }

  /**
   * Get {@link XParagraphCursor} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XParagraphCursor Interface.
   */
  public static XParagraphCursor XParagraphCursor(Object o)
  {
    return UnoRuntime.queryInterface(XParagraphCursor.class, o);
  }

  /**
   * Get {@link XServiceInfo} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XServiceInfo Interface.
   */
  public static XServiceInfo XServiceInfo(Object o)
  {
    return UnoRuntime.queryInterface(XServiceInfo.class, o);
  }

  /**
   * Get {@link XUpdatable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XUpdatable Interface.
   */
  public static XUpdatable XUpdatable(Object o)
  {
    return UnoRuntime.queryInterface(XUpdatable.class, o);
  }

  /**
   * Get {@link XTextViewCursorSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextViewCursorSupplier Interface.
   */
  public static XTextViewCursorSupplier XTextViewCursorSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextViewCursorSupplier.class, o);
  }

  /**
   * Get {@link XDrawPageSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDrawPageSupplier Interface.
   */
  public static XDrawPageSupplier XDrawPageSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XDrawPageSupplier.class, o);
  }

  /**
   * Get {@link XContentEnumerationAccess} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XContentEnumerationAccess Interface.
   */
  public static XContentEnumerationAccess XContentEnumerationAccess(Object o)
  {
    return UnoRuntime.queryInterface(XContentEnumerationAccess.class, o);
  }

  /**
   * Get {@link XControlShape} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XControlShape Interface.
   */
  public static XControlShape XControlShape(Object o)
  {
    return UnoRuntime.queryInterface(XControlShape.class, o);
  }

  /**
   * Get {@link XTextRange} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextRange Interface.
   */
  public static XTextRange XTextRange(Object o)
  {
    return UnoRuntime.queryInterface(XTextRange.class, o);
  }

  /**
   * Get {@link XDevice} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDevice Interface.
   */
  public static XDevice XDevice(Object o)
  {
    return UnoRuntime.queryInterface(XDevice.class, o);
  }

  /**
   * Get {@link XFramesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XFramesSupplier Interface.
   */
  public static XFramesSupplier XFramesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XFramesSupplier.class, o);
  }

  /**
   * Get {@link XDocumentPropertiesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDocumentPropertiesSupplier Interface.
   */
  public static XDocumentPropertiesSupplier XDocumentPropertiesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentPropertiesSupplier.class, o);
  }

  /**
   * Get {@link XDocumentProperties} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDocumentProperties Interface.
   */
  public static XDocumentProperties XDocumentProperties(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentProperties.class, o);
  }

  /**
   * Get {@link XDispatchProviderInterception} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDispatchProviderInterception Interface.
   */
  public static XDispatchProviderInterception XDispatchProviderInterception(Object o)
  {
    return UnoRuntime.queryInterface(XDispatchProviderInterception.class, o);
  }

  /**
   * Get {@link XTextCursor} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextCursor Interface.
   */
  public static XTextCursor XTextCursor(Object o)
  {
    return UnoRuntime.queryInterface(XTextCursor.class, o);
  }

  /**
   * Get {@link XPageCursor} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XPageCursor Interface.
   */
  public static XPageCursor XPageCursor(Object o)
  {
    return UnoRuntime.queryInterface(XPageCursor.class, o);
  }

  /**
   * Get {@link XNotifyingDispatch} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XNotifyingDispatch Interface.
   */
  public static XNotifyingDispatch XNotifyingDispatch(Object o)
  {
    return UnoRuntime.queryInterface(XNotifyingDispatch.class, o);
  }

  /**
   * Get {@link XStyleFamiliesSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStyleFamiliesSupplier Interface.
   */
  public static XStyleFamiliesSupplier XStyleFamiliesSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, o);
  }

  /**
   * Get {@link XStyle} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStyle Interface.
   */
  public static XStyle XStyle(Object o)
  {
    return UnoRuntime.queryInterface(XStyle.class, o);
  }

  /**
   * Get {@link XTextRangeCompare} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextRangeCompare Interface.
   */
  public static XTextRangeCompare XTextRangeCompare(Object o)
  {
    return UnoRuntime.queryInterface(XTextRangeCompare.class, o);
  }

  /**
   * Get {@link XTextSection} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextSection Interface.
   */
  public static XTextSection XTextSection(Object o)
  {
    return UnoRuntime.queryInterface(XTextSection.class, o);
  }

  /**
   * Get {@link XTextSectionsSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextSectionsSupplier Interface.
   */
  public static XTextSectionsSupplier XTextSectionsSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XTextSectionsSupplier.class, o);
  }

  /**
   * Get {@link XAcceleratorConfiguration} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XAcceleratorConfiguration Interface.
   */
  public static XAcceleratorConfiguration XAcceleratorConfiguration(Object o)
  {
    return UnoRuntime.queryInterface(XAcceleratorConfiguration.class, o);
  }

  /**
   * Get {@link XAccessible} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XAccessible Interface.
   */
  public static XAccessible XAccessible(Object o)
  {
    return UnoRuntime.queryInterface(XAccessible.class, o);
  }

  /**
   * Get {@link XUIConfigurationPersistence} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XUIConfigurationPersistence Interface.
   */
  public static XUIConfigurationPersistence XUIConfigurationPersistence(Object o)
  {
    return UnoRuntime.queryInterface(XUIConfigurationPersistence.class, o);
  }

  /**
   * Get {@link XModuleUIConfigurationManager} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XModuleUIConfigurationManager Interface.
   */
  public static XModuleUIConfigurationManager XModuleUIConfigurationManager(Object o)
  {
    return UnoRuntime.queryInterface(XModuleUIConfigurationManager.class, o);
  }

  /**
   * Get {@link XUIConfigurationManager} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XUIConfigurationManager Interface.
   */
  public static XUIConfigurationManager XUIConfigurationManager(Object o)
  {
    return UnoRuntime.queryInterface(XUIConfigurationManager.class, o);
  }

  /**
   * Get {@link XIndexContainer} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XIndexContainer Interface.
   */
  public static XIndexContainer XIndexContainer(Object o)
  {
    return UnoRuntime.queryInterface(XIndexContainer.class, o);
  }

  /**
   * Get {@link XFrame} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XFrame Interface.
   */
  public static XFrame XFrame(Object o)
  {
    return UnoRuntime.queryInterface(XFrame.class, o);
  }

  /**
   * Get {@link XTextColumns} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextColumns Interface.
   */
  public static XTextColumns XTextColumns(Object o)
  {
    return UnoRuntime.queryInterface(XTextColumns.class, o);
  }

  /**
   * Get {@link XStyleLoader} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStyleLoader Interface.
   */
  public static XStyleLoader XStyleLoader(Object o)
  {
    return UnoRuntime.queryInterface(XStyleLoader.class, o);
  }

  /**
   * Get {@link XPropertyState} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XPropertyState Interface.
   */
  public static XPropertyState XPropertyState(Object o)
  {
    return UnoRuntime.queryInterface(XPropertyState.class, o);
  }

  /**
   * Get {@link XStringSubstitution} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XStringSubstitution Interface.
   */
  public static XStringSubstitution XStringSubstitution(Object o)
  {
    return UnoRuntime.queryInterface(XStringSubstitution.class, o);
  }

  /**
   * Get {@link XRowSet} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XRowSet Interface.
   */
  public static XRowSet XRowSet(Object o)
  {
    return UnoRuntime.queryInterface(XRowSet.class, o);
  }

  /**
   * Get {@link XDocumentMetadataAccess} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDocumentMetadataAccess Interface.
   */
  public static XDocumentMetadataAccess XDocumentMetadataAccess(Object o)
  {
    return UnoRuntime.queryInterface(XDocumentMetadataAccess.class, o);
  }

  /**
   * Get {@link XRefreshable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XRefreshable Interface.
   */
  public static XRefreshable XRefreshable(Object o)
  {
    return UnoRuntime.queryInterface(XRefreshable.class, o);
  }
  
  /**
   * Get {@link XControl} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XControl Interface.
   */
  public static XControl XControl(Object object)
  {
    return UnoRuntime.queryInterface(XControl.class, object);
  }
  
  /**
   * Get {@link XFixedText} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XFixedText Interface.
   */
  public static XFixedText XFixedText(Object object)
  {
    return UnoRuntime.queryInterface(XFixedText.class, object);
  }
  
  /**
   * Get {@link XListBox} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XListBox Interface.
   */
  public static XListBox XListBox(Object object)
  {
    return UnoRuntime.queryInterface(XListBox.class, object);
  }
  
  /**
   * Get {@link XButton} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XButton Interface.
   */
  public static XButton XButton(Object object)
  {
    return UnoRuntime.queryInterface(XButton.class, object);
  }

  /**
   * Get {@link XCheckBox} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XCheckBox Interface.
   */
  public static XCheckBox XCheckBox(Object object)
  {
    return UnoRuntime.queryInterface(XCheckBox.class, object);
  }

  /**
   * Get {@link XRadioButton} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XRadioButton Interface.
   */
  public static XRadioButton XRadio(Object object)
  {
    return UnoRuntime.queryInterface(XRadioButton.class, object);
  }

  /**
   * Get {@link XNumericField} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XNumericField Interface.
   */
  public static XNumericField XNumericField(Object object)
  {
    return UnoRuntime.queryInterface(XNumericField.class, object);
  }

  /**
   * Get {@link XSpinField} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XSpinField Interface.
   */
  public static XSpinField XSpinField(Object object)
  {
    return UnoRuntime.queryInterface(XSpinField.class, object);
  }
  
  /**
   * Get {@link XTextComponent} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTextComponent Interface.
   */
  public static XTextComponent XTextComponent(Object o)
  {
    return UnoRuntime.queryInterface(XTextComponent.class, o);
  }
  
  /**
   * Get {@link XComboBox} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XComboBox Interface.
   */
  public static XComboBox XComboBox(Object o)
  {
    return UnoRuntime.queryInterface(XComboBox.class, o);
  }

  /**
   * Get {@link XContainerWindowProvider} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XContainerWindowProvider Interface.
   */
  public static XContainerWindowProvider XContainerWindowProvider(Object o)
  {
    return UnoRuntime.queryInterface(XContainerWindowProvider.class, o);
  }

  /**
   * Get {@link XDialog} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XDialog Interface.
   */
  public static XDialog XDialog(Object o)
  {
    return UnoRuntime.queryInterface(XDialog.class, o);
  }

  /**
   * Get {@link XCancellable} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XCancellable Interface.
   */
  public static XCancellable XCancellable(Object o)
  {
    return UnoRuntime.queryInterface(XCancellable.class, o);
  }

  /**
   * Get {@link XExtendedToolkit} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XExtendedToolkit Interface.
   */
  public static XExtendedToolkit XExtendedToolkit(Object o)
  {
    return UnoRuntime.queryInterface(XExtendedToolkit.class, o);
  }

  /**
   * Get {@link XIntrospection} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XIntrospection Interface.
   */
  public static XIntrospection XIntrospection(Object o)
  {
    return UnoRuntime.queryInterface(XIntrospection.class, o);
  }

  /**
   * Get {@link XMenu} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XMenu Interface.
   */
  public static XMenu XMenu(Object o)
  {
    return UnoRuntime.queryInterface(XMenu.class, o);
  }

  /**
   * Get {@link XItemList} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XItemList Interface.
   */
  public static XItemList XItemList(Object o)
  {
    return UnoRuntime.queryInterface(XItemList.class, o);
  }

  /**
   * Get {@link XPopupMenu} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XPopupMenu Interface.
   */
  public static XPopupMenu XPopupMenu(Object o)
  {
    return UnoRuntime.queryInterface(XPopupMenu.class, o);
  }

  /**
   * Get {@link XMutableTreeNode} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XMutableTreeNode Interface.
   */
  public static XMutableTreeNode XMutableTreeNode(Object o)
  {
    return UnoRuntime.queryInterface(XMutableTreeNode.class, o);
  }

  /**
   * Get {@link XModifiable2} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XModifiable2 Interface.
   */
  public static XModifiable2 XModifiable2(Object o)
  {
    return UnoRuntime.queryInterface(XModifiable2.class, o);
  }

  /**
   * Get {@link XSet} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XSet Interface.
   */
  public static XSet XSet(Object o)
  {
    return UnoRuntime.queryInterface(XSet.class, o);
  }

  /**
   * Get {@link XToolkit2} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XToolkit2 Interface.
   */
  public static XToolkit2 XToolkit2(Object o)
  {
    return UnoRuntime.queryInterface(XToolkit2.class, o);
  }

  /**
   * Get {@link XTreeControl} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XTreeControl Interface.
   */
  public static XTreeControl XTreeControl(Object o)
  {
    return UnoRuntime.queryInterface(XTreeControl.class, o);
  }

  /**
   * Get {@link XProgressBar} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XProgressBar Interface.
   */
  public static XProgressBar XProgressBar(Object o)
  {
    return UnoRuntime.queryInterface(XProgressBar.class, o);
  }

  /**
   * Liefert eine vorgeparste UNO-URL von urlStr.
   * 
   * @param urlStr
   * @return vorgeparste UNO-URL von urlStr.
   */
  public static com.sun.star.util.URL getParsedUNOUrl(String urlStr)
  {
    com.sun.star.util.URL[] unoURL = new com.sun.star.util.URL[]
    { new com.sun.star.util.URL() };
    unoURL[0].Complete = urlStr;
    if (urlTransformer != null)
    {
      urlTransformer.parseStrict(unoURL);
    }

    return unoURL[0];
  }

  /**
   * Liefert ein Service ConfigurationAccess mit dem der lesende Zugriff auf die
   * OOo-Configuration ab dem Knoten nodepath ermöglicht wird oder null, wenn der
   * Service nicht erzeugt werden kann.
   * 
   * @param nodepath
   *                   Beschreibung des Knotens des Konfigurationsbaumes, der als
   *                   neue Wurzel zurückgeliefert werden soll. Ein nodepath ist
   *                   z.B.
   *                   "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering"
   * @return ein ConfigurationUpdateAccess mit der Wurzel an dem Knoten nodepath
   *         oder null, falls der Service nicht erzeugt werden kann (wenn z.B. der
   *         Knoten nodepath nicht existiert).
   */
  public static XNameAccess getConfigurationAccess(String nodepath)
      throws UnoHelperException
  {
    PropertyValue[] props = new PropertyValue[]
    { new PropertyValue() };
    props[0].Name = "nodepath";
    props[0].Value = nodepath;
    Object confProv = getConfigurationProvider();
    try
    {
      return UNO.XNameAccess(UNO.XMultiServiceFactory(confProv)
          .createInstanceWithArguments(
              "com.sun.star.configuration.ConfigurationAccess", props));
    }
    catch (com.sun.star.uno.Exception e)
    {
      throw new UnoHelperException(e);
    }
  }

  /**
   * Liefert ein Service ConfigurationUpdateAccess mit dem der lesende und
   * schreibende Zugriff auf die OOo-Configuration ab dem Knoten nodepath
   * ermöglicht wird oder null wenn der Service nicht erzeugt werden kann.
   * 
   * @param nodepath
   *                   Beschreibung des Knotens des Konfigurationsbaumes, der als
   *                   neue Wurzel zurückgeliefert werden soll. Ein nodepath ist
   *                   z.B.
   *                   "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering"
   * @return ein ConfigurationUpdateAccess mit der Wurzel an dem Knoten nodepath
   *         oder null, falls der Service nicht erzeugt werden kann (wenn z.B. der
   *         Knoten nodepath nicht existiert).
   * @throws UnoHelperException 
   */
  public static XChangesBatch getConfigurationUpdateAccess(String nodepath)
      throws UnoHelperException
  {
    PropertyValue[] props = new PropertyValue[]
    { new PropertyValue() };
    props[0].Name = "nodepath";
    props[0].Value = nodepath;
    Object confProv = getConfigurationProvider();
    try
    {
      return UNO.XChangesBatch(UnoService.createServiceWithArguments(
          UnoService.CSS_CONFIGURATION_CONFIGURATION_UPDATE_ACCESS, props, confProv));
    } catch (Exception e)
    {
      throw new UnoHelperException(e);
    }
  }

  /**
   * Liefert den configurationProvider, mit dem der Zugriff auf die Konfiguration
   * von OOo ermöglicht wird.
   * 
   * @return ein neuer configurationProvider
   */
  private static Object getConfigurationProvider()
  {
    if (configurationProvider == null)
    {
      configurationProvider = UnoComponent
          .createComponentWithContext(UnoComponent.CSS_CONFIGURATION_CONFIGURATION_PROVIDER);
    }
    return configurationProvider;
  }

  /**
   * Liefert den shortcutManager zu der OOo Komponente component zurück.
   * 
   * @param component
   *                    die OOo Komponente zu der der ShortcutManager geliefert
   *                    werden soll z.B "com.sun.star.text.TextDocument"
   * @return der shortcutManager zur OOo Komponente component oder null falls kein
   *         shortcutManager erzeugt werden kann.
   * @throws UnoHelperException 
   * 
   */
  public static XAcceleratorConfiguration getShortcutManager(String component)
      throws UnoHelperException
  {
    // XModuleUIConfigurationManagerSupplier moduleUICfgMgrSupplier
    XModuleUIConfigurationManagerSupplier moduleUICfgMgrSupplier = UNO.XModuleUIConfigurationManagerSupplier(
        UnoComponent.createComponentWithContext(UnoComponent.CSS_UI_MODULE_UI_CONFIGURATION_MANAGER_SUPPLIER));

    if (moduleUICfgMgrSupplier == null)
    {
      return null;
    }

    try
    {
      // XCUIConfigurationManager moduleUICfgMgr
      XUIConfigurationManager moduleUICfgMgr = null;

      moduleUICfgMgr = moduleUICfgMgrSupplier.getUIConfigurationManager(component);

      // XAcceleratorConfiguration xAcceleratorConfiguration
      Method m = moduleUICfgMgr.getClass().getMethod("getShortCutManager", (Class[]) null);
      return UNO.XAcceleratorConfiguration(m.invoke(moduleUICfgMgr, (Object[]) null));
    }
    catch (NoSuchElementException | NoSuchMethodException | SecurityException
        | IllegalAccessException | java.lang.IllegalArgumentException
        | InvocationTargetException e)
    {
      throw new UnoHelperException("Kein Zugriff auf den ShrotcutManager.", e);
    }
  }

  /**
   * Wenn hide=true ist, so wird die Eigenschaft CharHidden für range auf true
   * gesetzt und andernfalls der Standardwert (=false) für die Property CharHidden
   * wieder hergestellt. Dadurch lässt sich der Text in range unsichtbar schalten
   * bzw. wieder sichtbar schalten. Die Repräsentation von unsichtbar geschaltenen
   * Stellen erfolgt in der Art, dass OOo für den unsichtbaren Textbereich ein
   * neuen automatisch generierten Character-Style anlegt, der die Eigenschaften
   * der bisher gesetzten Styles erbt und lediglich die Eigenschaft "Sichtbarkeit"
   * auf unsichtbar setzt. Beim Aufheben einer unsichtbaren Stelle sorgt das
   * Zurücksetzen auf den Standardwert dafür, dass der vorher angelegte
   * automatische-Style wieder zurück genommen wird - so ist sichergestellt, dass
   * das Aus- und Wiedereinblenden von Textbereichen keine Änderungen der bisher
   * gesetzten Styles hervorruft.
   * 
   * @param range
   *                Der Textbereich, der aus- bzw. eingeblendet werden soll.
   * @param hide
   *                hide=true blendet aus, hide=false blendet ein.
   * @throws UnoHelperException 
   */
  public static void hideTextRange(XTextRange range, boolean hide)
      throws UnoHelperException
  {
    String propName = "CharHidden";
    if (hide)
    {
      UnoProperty.setProperty(range, propName, Boolean.TRUE);
      // Workaround für update Bug
      // http://qa.openoffice.org/issues/show_bug.cgi?id=78896
      UnoProperty.setProperty(range, propName, Boolean.FALSE);
      UnoProperty.setProperty(range, propName, Boolean.TRUE);
    } else
    {
      // Workaround für (den anderen) update Bug
      // http://qa.openoffice.org/issues/show_bug.cgi?id=103101
      // Nur das Rücksetzen auf den Standardwert reicht nicht aus. Daher erfolgt
      // vor
      // dem Rücksetzen auf den Standardwert eine explizite Einblendung.
      UnoProperty.setProperty(range, propName, Boolean.FALSE);
      UnoProperty.setPropertyToDefault(range, propName);
    }
  }

  /**
   * Iteriert über alle Paragraphen in einer XTextRange.
   * 
   * @param range
   * @param c
   *                Lambda-Funktion, die den aktuellen Paragraphen als Parameter
   *                erhält.
   */
  public static void forEachParagraphInRange(XTextRange range, Consumer<Object> c)
  {
    XTextCursor cursor = range.getText().createTextCursorByRange(range);

    for (XEnumerationAccess par : UnoCollection.getCollection(cursor,
        XEnumerationAccess.class))
    {
      if (par != null)
      {
        c.accept(par);
      }
    }
  }

  /**
   * Iteriert über alle Text-Objekte in einer XTextRange.
   * 
   * @param range
   * @param c
   *                Lambda-Funktion, die das aktuelle Text-Objekt als Parameter
   *                erhält.
   */
  public static void forEachTextPortionInRange(XTextRange range, Consumer<Object> c)
  {
    XTextCursor cursor = range.getText().createTextCursorByRange(range);

    for (XEnumerationAccess parEnum : UnoCollection.getCollection(cursor,
        XEnumerationAccess.class))
    {
      if (parEnum != null)
      {
        for (Object o : UnoCollection.getCollection(parEnum, Object.class))
        {
          c.accept(o);
        }
      }
    }
  }

  /**
   * Get the first element of the UNO object which have to implement {@link XEnumerationAccess}.
   * 
   * @param o
   *          The UNO object.
   * @return The first element.
   */
  public static Object getFirstElementInEnumeration(Object o)
  {
    UnoIterator<Object> iter = UnoIterator.create(o, Object.class);
    if (iter.hasNext())
    {
      return iter.next();
    }

    return null;
  }
}

package de.muenchen.allg.afid;

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
import com.sun.star.awt.XLayoutConstrains;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMenu;
import com.sun.star.awt.XNumericField;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.awt.XProgressBar;
import com.sun.star.awt.XRadioButton;
import com.sun.star.awt.XScrollBar;
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
import com.sun.star.awt.tab.XTabPageModel;
import com.sun.star.awt.tree.XMutableTreeDataModel;
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
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          If true macros can be executed.
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, boolean allowMacros)
      throws UnoHelperException
  {
      return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          If true macros can be executed.
   * @param hidden
   *          If true they document is opened invisible.
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
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
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          The macro execution mode (@link {@link MacroExecMode}).
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, short allowMacros)
      throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, false);
  }

  /**
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          The macro execution mode (@link {@link MacroExecMode}).
   * @param hidden
   *          If true they document is opened invisible.
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, short allowMacros,
      boolean hidden) throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate, allowMacros, hidden, new PropertyValue[0]);
  }

  /**
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          The macro execution mode (@link {@link MacroExecMode}).
   * @param args
   *          Additional arguments for loading the document.
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, boolean allowMacros,
      PropertyValue... args) throws UnoHelperException
  {
    return loadComponentFromURL(url, asTemplate,
        (allowMacros) ? MacroExecMode.ALWAYS_EXECUTE_NO_WARN : MacroExecMode.NEVER_EXECUTE, false, args);
  }

  /**
   * Load a document and initialize {@link #compo}.
   *
   * @param url
   *          The URL of the docment (e.g. "file:///C:/temp/footest.odt" or
   *          "private:factory/swriter").
   * @param asTemplate
   *          If true the document is treated as a template and a new unnamed document is created
   *          based on the template.
   * @param allowMacros
   *          The macro execution mode (@link {@link MacroExecMode}).
   * @param hidden
   *          If true they document is opened invisible.
   * @param args
   *          Additional arguments for loading the document.
   * @return The document.
   * @throws UnoHelperException
   *           Can't load the document.
   */
  public static XComponent loadComponentFromURL(String url, boolean asTemplate, short allowMacros, boolean hidden,
      PropertyValue... args) throws UnoHelperException
  {
    try
    {
      XComponentLoader loader = UNO.XComponentLoader(desktop);
      UnoProps props = new UnoProps(args);
      props.setPropertyValue(UnoProperty.MACRO_EXECUTION_MODE, allowMacros);
      props.setPropertyValue(UnoProperty.AS_TEMPLATE, asTemplate);
      props.setPropertyValue(UnoProperty.HIDDEN, hidden);
      PropertyValue[] arguments = props.getProps();
      arguments = ArrayUtils.addAll(arguments, args);

      XComponent lc = loader.loadComponentFromURL(url, "_blank", FrameSearchFlag.CREATE, arguments);
      if (lc != null)
      {
        compo = lc;
      }
      return lc;
    }
    catch (IllegalArgumentException | IOException e)
    {
      throw new UnoHelperException("Can't load the document", e);
    }
  }
  
  /**
   * Convert file path in a system specific URL.
   *
   * @param filePath
   *          The path to convert.
   * @return A valid URL for Office an empty String with this isn't possible.
   * @throws UnoHelperException
   *           Can't convert the path.
   */
  public static String convertFilePathToURL(String filePath) throws UnoHelperException
  {
    String resultURL = null;

    try
    {
      XFileIdentifierConverter xFileConverter = UNO.XFileIdentifierConverter(
          UnoComponent.createComponentWithContext(UnoComponent.CSS_UCB_FILE_CONTENT_PROVIDER));
      resultURL = xFileConverter.getFileURLFromSystemPath("", filePath);
    } catch (Exception e)
    {
      throw new UnoHelperException("", e);
    }

    return resultURL;
  }

  /**
   * Get the the document which has the focus.
   *
   * @return The text document which has the focus or null if it isn't a text document.
   */
  public static XTextDocument getCurrentTextDocument()
  {
    XComponent xComponent = desktop.getCurrentComponent();

    return UNO.XTextDocument(xComponent);
  }

  /**
   * Dispatch a command on a model.
   *
   * @param doc
   *          The model.
   * @param url
   *          The command as URL.
   */
  public static void dispatch(XModel doc, String url)
  {
    XDispatchProvider prov = UNO.XDispatchProvider(doc.getCurrentController().getFrame());
    dispatchHelper.executeDispatch(prov, url, "", FrameSearchFlag.SELF, new PropertyValue[] {});
  }

  /**
   * Dispatch a command on a text document and wait until execution finished.
   *
   * @param doc
   *          The text document.
   * @param url
   *          The command URL.
   * @return The dispatch result or null if no command exists or the dispatch is aborted.
   */
  public static DispatchResultEvent dispatchAndWait(XTextDocument doc, String url)
  {
    if (doc == null)
    {
      return null;
    }

    URL unoUrl = getParsedUNOUrl(url);

    XDispatchProvider prov = UNO.XDispatchProvider(doc.getCurrentController().getFrame());
    if (prov == null)
    {
      return null;
    }

    XNotifyingDispatch disp = UNO.XNotifyingDispatch(prov.queryDispatch(unoUrl, "", FrameSearchFlag.SELF));
    if (disp == null)
    {
      return null;
    }

    final boolean[] lock = new boolean[] { true };
    final DispatchResultEvent[] resultEvent = new DispatchResultEvent[] { null };

    disp.dispatchWithNotification(unoUrl, new PropertyValue[] {}, new XDispatchResultListener()
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
          Thread.currentThread().interrupt();
        }
    }
    return resultEvent[0];
  }

  /**
   * Call a macro.
   *
   * @param scriptProviderOrSupplier
   *          The provider of the macro (e.g. a text document).
   * @param macroName
   *          The name of the macro. Library and module are separate by a dot (e.g.
   *          "Standard.Module1.Foo"). But "Foo" is also valid.
   * @param args
   *          The arguments for the macro.
   * @param location
   *          List of possible locations (e.g. "application", "share", "document"). The order is
   *          important if there are more macros with the same name but in different locations.
   * @throws RuntimeException
   *           If a macro with this name can't be found.
   * @return The return value of the macro.
   * @throws UnoHelperException
   *           Can't access the macros.
   */
  public static Object executeMacro(Object scriptProviderOrSupplier, String macroName, Object[] args,
      String[] location) throws UnoHelperException
  {
    XScriptProvider provider = UNO.XScriptProvider(scriptProviderOrSupplier);
    if (provider == null)
    {
      XScriptProviderSupplier supp = UNO.XScriptProviderSupplier(scriptProviderOrSupplier);
      if (supp == null)
      {
        throw new RuntimeException("No script provider or supplier given.");
      }
      provider = supp.getScriptProvider();
    }

    XBrowseNode root = UNO.XBrowseNode(provider);
    return Utils.executeMacroInternal(macroName, args, null, root, location);
  }

  /**
   * Call a macro not stored in a document. Macros in location "application" have priority.
   *
   * @param macroName
   *          The name of the macro. Library and module are separate by a dot (e.g.
   *          "Standard.Module1.Foo"). But "Foo" is also valid.
   * @param args
   *          The arguments for the macro.
   * @throws RuntimeException
   *           If a macro with this name can't be found.
   * @return The return value of the macro.
   * @throws UnoHelperException
   *           Can't access the macros.
   */
  public static Object executeGlobalMacro(String macroName, Object[] args) throws UnoHelperException
  {
    final String[] userAndShare = new String[] { "application", "share" };
    return Utils.executeMacroInternal(macroName, args, null, scriptRoot.unwrap(), userAndShare);
  }

  /**
   * Call a macro.
   *
   * @param macroName
   *          The name of the macro. Library and module are separate by a dot (e.g.
   *          "Standard.Module1.Foo"). But "Foo" is also valid.
   * @param args
   *          The arguments for the macro.
   * @param location
   *          List of possible locations (e.g. "application", "share", "document"). The order is
   *          important if there are more macros with the same name but in different locations.
   * @throws RuntimeException
   *           If a macro with this name can't be found.
   * @return The return value of the macro.
   * @throws UnoHelperException
   *           Can't access the macros.
   */
  public static Object executeMacro(String macroName, Object[] args, String[] location) throws UnoHelperException
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
   * Search for a script in locations document, application, share.
   *
   * @param xBrowseNode
   *          Root of the tree to scan.
   * @param prefix
   *          Prefix of the script to find.
   * @param nameToFind
   *          Script to find. Path can be separated by dots.
   * @param caseSensitive
   *          If true search is case sensitive.
   * @return The node of the script or null if the script can't be found.
   * @throws UnoHelperException
   *           Can't find the script.
   */
  public static XBrowseNode findBrowseNodeTreeLeaf(XBrowseNode xBrowseNode, String prefix, String nameToFind,
      boolean caseSensitive) throws UnoHelperException
  {
    XBrowseNodeAndXScriptProvider x = findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind,
        caseSensitive);
    return x.getXBrowseNode();
  }

  /**
   * Search for a script in locations document, application, share.
   *
   * @param xBrowseNode
   *          Root of the tree to scan.
   * @param prefix
   *          Prefix of the script to find.
   * @param nameToFind
   *          Script to find. Path can be separated by dots.
   * @param caseSensitive
   *          If true search is case sensitive.
   * @return The node of the script and the first parent which implements {@link XScriptProvider}.
   *         If one of them isn't found its value is <code>null</code>.
   * @throws UnoHelperException
   *           Can't find the script.
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode,
      String prefix, String nameToFind, boolean caseSensitive) throws UnoHelperException
  {
    return findBrowseNodeTreeLeafAndScriptProvider(xBrowseNode, prefix, nameToFind, caseSensitive, null);
  }

  /**
   * Search for a script.
   *
   * @param xBrowseNode
   *          Root of the tree to scan.
   * @param prefix
   *          Prefix of the script to find.
   * @param nameToFind
   *          Script to find. Path can be separated by dots.
   * @param caseSensitive
   *          If true search is case sensitive.
   * @param location
   *          List of possible locations (e.g. "application", "share", "document"). The order is
   *          important if there are more macros with the same name but in different locations. If
   *          <code>location==null</code> <code>{"document", "application", "share"}</code> is used.
   * @return The node of the script and the first parent which implements {@link XScriptProvider}.
   *         If one of them isn't found its value is <code>null</code>.
   * @throws UnoHelperException
   *           Can't find the script.
   */
  public static XBrowseNodeAndXScriptProvider findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode xBrowseNode,
      String prefix, String nameToFind, boolean caseSensitive, String[] location) throws UnoHelperException
  {
    String[] loc;
    if (location == null)
    {
      loc = new String[] { "document", "application", "share" };
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
   * Get {@link XScriptProviderSupplier} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XScriptProviderSupplier} Interface.
   */
  public static XScriptProviderSupplier XScriptProviderSupplier(Object o)
  {
    return UnoRuntime.queryInterface(XScriptProviderSupplier.class, o);
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
   * Get {@link XTabPageModel} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns {@link XTabPageModel} Interface.
   */
  public static XTabPageModel XTabPageModel(Object o)
  {
    return UnoRuntime.queryInterface(XTabPageModel.class, o);
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
   * Get {@link XScrollBar} Interface from Object.
   *
   * @param object
   *          Object.
   * @return Returns XScrollBar Interface.
   */
  public static XScrollBar XScrollBar(Object object)
  {
    return UnoRuntime.queryInterface(XScrollBar.class, object);
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
   * Get {@link XMutableTreeDataModel} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XMutableTreeDataModel Interface.
   */
  public static XMutableTreeDataModel XMutableTreeDataModel(Object o)
  {
    return UnoRuntime.queryInterface(XMutableTreeDataModel.class, o);
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
   * Get {@link XFileIdentifierConverter} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XFileIdentifierConverter Interface.
   */
  public static XFileIdentifierConverter XFileIdentifierConverter(Object o)
  {
    return UnoRuntime.queryInterface(XFileIdentifierConverter.class, o);
  }

  /**
   * Get {@link XLayoutConstrains} Interface from Object.
   *
   * @param o
   *          Object.
   * @return Returns XLayoutConstrains Interface.
   */
  public static XLayoutConstrains XLayoutConstrains(Object o)
  {
    return UnoRuntime.queryInterface(XLayoutConstrains.class, o);
  }

  /**
   * Parse an UNO-URL.
   * 
   * @param urlStr
   *          The UNO-URL.
   * @return parsed UNO-URL.
   */
  public static com.sun.star.util.URL getParsedUNOUrl(String urlStr)
  {
    com.sun.star.util.URL[] unoURL = new com.sun.star.util.URL[] { new com.sun.star.util.URL() };
    unoURL[0].Complete = urlStr;
    if (urlTransformer != null)
    {
      urlTransformer.parseStrict(unoURL);
    }

    return unoURL[0];
  }

  /**
   * Get access a part of the configuration so that it can be read.
   *
   * @param nodepath
   *          The part of the configuration (e.g.
   *          "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering")
   * @return The updatable configuration part or null if the part doesn't exist.
   */
  public static XNameAccess getConfigurationAccess(String nodepath)
  {
    UnoProps props = new UnoProps(UnoProperty.NODEPATH, nodepath);
    Object confProv = getConfigurationProvider();
    return UNO.XNameAccess(UnoService.createServiceWithArguments(UnoService.CSS_CONFIGURATION_CONFIGURATION_ACCESS,
        props.getProps(), confProv));
  }

  /**
   * Get access a part of the configuration so that it can be read and written.
   *
   * @param nodepath
   *          The part of the configuration (e.g.
   *          "/org.openoffice.Office.Writer/AutoFunction/Format/ByInput/ApplyNumbering")
   * @return The updatable configuration part or null if the part doesn't exist.
   * @throws UnoHelperException
   *           Can't get access to a updatable configuration.
   */
  public static XChangesBatch getConfigurationUpdateAccess(String nodepath) throws UnoHelperException
  {
    PropertyValue[] props = new PropertyValue[] { new PropertyValue() };
    props[0].Name = "nodepath";
    props[0].Value = nodepath;
    Object confProv = getConfigurationProvider();
    try
    {
      return UNO.XChangesBatch(UnoService
          .createServiceWithArguments(UnoService.CSS_CONFIGURATION_CONFIGURATION_UPDATE_ACCESS, props, confProv));
    } catch (Exception e)
    {
      throw new UnoHelperException(e);
    }
  }

  /**
   * Get access to the configuration of Office.
   *
   * @return A configuration provider.
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
   * Get the accelerator configuration.
   *
   * @param component
   *          The FQDN of the component for which the accelerator configuration is requested (e.g.
   *          "com.sun.star.text.TextDocument")
   * @return The accelerator configuration or null if there isn't a accelerator configuration.
   * @throws UnoHelperException
   *           Can't access the accelerator configuration.
   */
  public static XAcceleratorConfiguration getShortcutManager(String component) throws UnoHelperException
  {
    XModuleUIConfigurationManagerSupplier moduleUICfgMgrSupplier = UNO.XModuleUIConfigurationManagerSupplier(
        UnoComponent.createComponentWithContext(UnoComponent.CSS_UI_MODULE_UI_CONFIGURATION_MANAGER_SUPPLIER));

    if (moduleUICfgMgrSupplier == null)
    {
      return null;
    }

    try
    {
      XUIConfigurationManager moduleUICfgMgr = moduleUICfgMgrSupplier.getUIConfigurationManager(component);
      return moduleUICfgMgr.getShortCutManager();
    } catch (NoSuchElementException | SecurityException | java.lang.IllegalArgumentException e)
    {
      throw new UnoHelperException("No shortcut manager", e);
    }
  }

  /**
   * Hide or show a text range.
   *
   * @param range
   *          The text range to hide or show.
   * @param hide
   *          True if the text range should be invisible, false if visible.
   * @throws UnoHelperException
   *           Can't modify the property.
   */
  public static void hideTextRange(XTextRange range, boolean hide) throws UnoHelperException
  {
    if (hide)
    {
      UnoProperty.setProperty(range, UnoProperty.CHAR_HIDDEN, Boolean.TRUE);
      // Workaround for http://qa.openoffice.org/issues/show_bug.cgi?id=78896
      UnoProperty.setProperty(range, UnoProperty.CHAR_HIDDEN, Boolean.FALSE);
      UnoProperty.setProperty(range, UnoProperty.CHAR_HIDDEN, Boolean.TRUE);
    } else
    {
      // Workaround for http://qa.openoffice.org/issues/show_bug.cgi?id=103101
      UnoProperty.setProperty(range, UnoProperty.CHAR_HIDDEN, Boolean.FALSE);
      UnoProperty.setPropertyToDefault(range, UnoProperty.CHAR_HIDDEN);
    }
  }

  /**
   * Perform an action on all paragraphs in a text range.
   *
   * @param range
   *          The text range.
   * @param consumer
   *          The action to perform on the paragraphs.
   */
  public static void forEachParagraphInRange(XTextRange range, Consumer<Object> consumer)
  {
    XTextCursor cursor = range.getText().createTextCursorByRange(range);

    for (XEnumerationAccess par : UnoCollection.getCollection(cursor, XEnumerationAccess.class))
    {
      if (par != null)
      {
        consumer.accept(par);
      }
    }
  }

  /**
   * Perform an action on all text objects in a text range.
   *
   * @param range
   *          The text range.
   * @param consumer
   *          The action to perform on the text objects.
   */
  public static void forEachTextPortionInRange(XTextRange range, Consumer<Object> consumer)
  {
    XTextCursor cursor = range.getText().createTextCursorByRange(range);

    for (XEnumerationAccess parEnum : UnoCollection.getCollection(cursor, XEnumerationAccess.class))
    {
      if (parEnum != null)
      {
        for (Object o : UnoCollection.getCollection(parEnum, Object.class))
        {
          consumer.accept(o);
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

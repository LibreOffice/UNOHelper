/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 The Document Foundation
 * %%
 * Licensed under the EUPL, Version 1.1 or – as soon they will be
 * approved by the European Commission - subsequent versions of the
 * EUPL (the "Licence");
 *
 * You may not use this work except in compliance with the Licence.
 * You may obtain a copy of the Licence at:
 *
 * http://ec.europa.eu/idabc/eupl5
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the Licence is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the Licence for the specific language governing permissions and
 * limitations under the Licence.
 * #L%
 */
package org.libreoffice.ext.unohelper.util;

import org.libreoffice.ext.unohelper.common.UNO;
import org.libreoffice.ext.unohelper.common.UnoHelperRuntimeException;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;

/**
 * Helper for creating UNO components.
 */
@SuppressWarnings("java:S103")
public final class UnoComponent
{

  public static final String CSS_AWT_CONTAINER_WINDOW_PROVIDER = "com.sun.star.awt.ContainerWindowProvider";
  public static final String CSS_AWT_POPUP_MENU = "com.sun.star.awt.PopupMenu";
  public static final String CSS_AWT_TAB_UNO_CONTROL_TAB_PAGE = "com.sun.star.awt.tab.UnoControlTabPage";
  public static final String CSS_AWT_TAB_UNO_CONTROL_TAB_PAGE_CONTAINER = "com.sun.star.awt.tab.UnoControlTabPageContainer";
  public static final String CSS_AWT_TAB_UNO_CONTROL_TAB_PAGE_CONTAINER_MODEL = "com.sun.star.awt.tab.UnoControlTabPageContainerModel";
  public static final String CSS_AWT_TOOLKIT = "com.sun.star.awt.Toolkit";
  public static final String CSS_AWT_TREE_MUTABLE_TREE_DATA_MODEL = "com.sun.star.awt.tree.MutableTreeDataModel";
  public static final String CSS_AWT_TREE_TREE_CONTROL = "com.sun.star.awt.tree.TreeControl";
  public static final String CSS_AWT_UNO_CONTROL_BUTTON = "com.sun.star.awt.UnoControlButton";
  public static final String CSS_AWT_UNO_CONTROL_CHECK_BOX = "com.sun.star.awt.UnoControlCheckBox";
  public static final String CSS_AWT_UNO_CONTROL_COMBO_BOX = "com.sun.star.awt.UnoControlComboBox";
  public static final String CSS_AWT_UNO_CONTROL_CONTAINER = "com.sun.star.awt.UnoControlContainer";
  public static final String CSS_AWT_UNO_CONTROL_CONTAINER_MODEL = "com.sun.star.awt.UnoControlContainerModel";
  public static final String CSS_AWT_UNO_CONTROL_DIALOG = "com.sun.star.awt.UnoControlDialog";
  public static final String CSS_AWT_UNO_CONTROL_EDIT = "com.sun.star.awt.UnoControlEdit";
  public static final String CSS_AWT_UNO_CONTROL_FIXED_HYPER_LINK = "com.sun.star.awt.UnoControlFixedHyperlink";
  public static final String CSS_AWT_UNO_CONTROL_FIXED_LINE = "com.sun.star.awt.UnoControlFixedLine";
  public static final String CSS_AWT_UNO_CONTROL_FIXED_TEXT = "com.sun.star.awt.UnoControlFixedText";
  public static final String CSS_AWT_UNO_CONTROL_LIST_BOX = "com.sun.star.awt.UnoControlListBox";
  public static final String CSS_AWT_UNO_CONTROL_NUMERIC_FIELD = "com.sun.star.awt.UnoControlNumericField";
  public static final String CSS_AWT_UNO_CONTROL_SCROLL_BAR = "com.sun.star.awt.UnoControlScrollBar";
  public static final String CSS_BRIDGE_UNO_URL_RESOLVER = "com.sun.star.bridge.UnoUrlResolver";
  public static final String CSS_CONFIGURATION_CONFIGURATION_PROVIDER = "com.sun.star.configuration.ConfigurationProvider";
  public static final String CSS_FRAME_DESKTOP = "com.sun.star.frame.Desktop";
  public static final String CSS_FRAME_DISPATCH_HELPER = "com.sun.star.frame.DispatchHelper";
  public static final String CSS_SDB_DATABASE_CONTEXT = "com.sun.star.sdb.DatabaseContext";
  public static final String CSS_SDB_ROW_SET = "com.sun.star.sdb.RowSet";
  public static final String CSS_TEXT_AUTO_TEXT_CONTAINER = "com.sun.star.text.AutoTextContainer";
  public static final String CSS_UCB_FILE_CONTENT_PROVIDER = "com.sun.star.ucb.FileContentProvider";
  public static final String CSS_UI_MODULE_UI_CONFIGURATION_MANAGER_SUPPLIER = "com.sun.star.ui.ModuleUIConfigurationManagerSupplier";
  public static final String CSS_UI_UI_ELEMENT_FACTORY_MANAGER = "com.sun.star.ui.UIElementFactoryManager";
  public static final String CSS_UTIL_PATH_SETTINGS = "com.sun.star.util.PathSettings";
  public static final String CSS_UTIL_PATH_SUBSTITUTION = "com.sun.star.util.PathSubstitution";
  public static final String CSS_UTIL_URL_TRANSFORMER = "com.sun.star.util.URLTransformer";

  /**
   * Calls {@link #createComponentWithContext(String, XMultiComponentFactory, XComponentContext)
   * with {@link UNO#xMCF} and {@link UNO#defaultContext}.
   *
   * @param componentName
   *          The name of the component.
   * @return The component.
   * @throws UnoHelperRuntimeException
   *           Component can't be created.
   */
  public static Object createComponentWithContext(String componentName)
  {
    return createComponentWithContext(componentName, UNO.xMCF, UNO.defaultContext);
  }
  /**
   * Create a component.
   *
   * @param componentName
   *          The name of the component.
   * @param factory
   *          The service manager.
   * @param context
   *          The context of the component.
   * @return The component.
   * @throws UnoHelperRuntimeException
   *           Component can't be created.
   */
  public static Object createComponentWithContext(String componentName,
      XMultiComponentFactory factory, XComponentContext context)
  {
    try
    {
      return factory.createInstanceWithContext(componentName, context);
    } catch (com.sun.star.uno.Exception e)
    {
      throw new UnoHelperRuntimeException("UNO-Service konnte nicht erstellt werden.", e);
    }
  }

  private UnoComponent()
  {
  }
}

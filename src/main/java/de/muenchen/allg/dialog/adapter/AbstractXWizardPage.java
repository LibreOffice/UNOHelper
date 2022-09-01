/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2022 Landeshauptstadt München
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
package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XContainerWindowProvider;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.lang.XEventListener;
import com.sun.star.ui.dialogs.XWizardPage;
import com.sun.star.uno.Exception;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.util.UnoComponent;

/**
 * Provides default implementations of standard methods for the {@link XWizardPage}.
 */
public abstract class AbstractXWizardPage implements XWizardPage
{

  /**
   * The id of this page.
   */
  protected short pageId;

  /**
   * The window of the page.
   */
  protected XWindow window;

  /**
   * Create a wizard page with dialog.
   *
   * @param pageId
   *          The id of the page.
   * @param parentWindow
   *          The parent window of the page.
   * @param dialogDescritpion
   *          Path to an dialog description.
   * @throws Exception
   *           The dialog can't be created.
   */
  public AbstractXWizardPage(short pageId, XWindow parentWindow, String dialogDescritpion) throws Exception
  {
    this.pageId = pageId;

    XWindowPeer peer = UNO.XWindowPeer(parentWindow);
    XContainerWindowProvider provider;

    provider = UNO.XContainerWindowProvider(
        UnoComponent.createComponentWithContext(UnoComponent.CSS_AWT_CONTAINER_WINDOW_PROVIDER));
    window = provider.createContainerWindow(dialogDescritpion, "", peer, null);
  }

  @Override
  public void addEventListener(XEventListener arg0)
  {
    // default implementation
  }

  @Override
  public void dispose()
  {
    window.dispose();
  }

  @Override
  public void removeEventListener(XEventListener arg0)
  {
    // default implementation
  }

  @Override
  public void activatePage()
  {
    window.setVisible(true);
  }

  @Override
  public boolean canAdvance()
  {
    return false;
  }

  @Override
  public boolean commitPage(short arg0)
  {
    return false;
  }

  @Override
  public short getPageId()
  {
    return pageId;
  }

  @Override
  public XWindow getWindow()
  {
    return window;
  }

}

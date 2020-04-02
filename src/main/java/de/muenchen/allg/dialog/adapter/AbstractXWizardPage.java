package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XContainerWindowProvider;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.lang.XEventListener;
import com.sun.star.ui.dialogs.XWizardPage;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

import de.muenchen.allg.afid.UNO;

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

    provider = UnoRuntime.queryInterface(XContainerWindowProvider.class, UNO.xMCF
        .createInstanceWithContext("com.sun.star.awt.ContainerWindowProvider", UNO.defaultContext));
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

package de.muenchen.allg.util;

import com.sun.star.frame.XController2;
import com.sun.star.ui.XDeck;
import com.sun.star.ui.XDecks;
import com.sun.star.ui.XSidebarProvider;

import de.muenchen.allg.afid.UnoDictionary;
import de.muenchen.allg.afid.UnoHelperException;

/**
 * Helper for accessing Sidebars in LibreOffice.
 */
public class UnoSidebar
{

  private UnoSidebar()
  {
  }

  /**
   * Get all decks known by LibreOffice.
   *
   * @param xController
   *          The controller of the document.
   * @return returns The decks or NULL if XSidebarProvider returns no decks.
   * @throws UnoHelperException
   *           Can't access the sidebar.
   */
  private static XDecks getDecks(XController2 xController) throws UnoHelperException
  {
    XSidebarProvider sideBarProvider = xController.getSidebar();

    if (sideBarProvider != null)
    {
      return sideBarProvider.getDecks();
    }
    throw new UnoHelperException("Can't access the sidebar.");
  }

  /**
   * Get a deck with a specific name.
   *
   * @param name
   *          The name of the deck.
   * @param xController
   *          The controller of the document.
   * @return The deck or null if it can't be found.
   * @throws UnoHelperException
   *           Can't access the sidebar.
   */
  public static XDeck getDeckByName(String name, XController2 xController) throws UnoHelperException
  {
    UnoDictionary<XDeck> xDecks = UnoDictionary.create(getDecks(xController), XDeck.class);
    if (xDecks != null)
    {
      return xDecks.get(name);
    }
    return null;
  }

}

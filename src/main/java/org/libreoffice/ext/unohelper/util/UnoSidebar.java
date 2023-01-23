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

import org.libreoffice.ext.unohelper.common.UnoDictionary;
import org.libreoffice.ext.unohelper.common.UnoHelperException;

import com.sun.star.frame.XController2;
import com.sun.star.ui.XDeck;
import com.sun.star.ui.XDecks;
import com.sun.star.ui.XSidebarProvider;

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

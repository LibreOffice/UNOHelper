/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 Landeshauptstadt München
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

import com.sun.star.awt.MenuEvent;
import com.sun.star.awt.XMenuListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XMenuListener}.
 */
public abstract class AbstractMenuListener implements XMenuListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void itemActivated(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemDeactivated(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemHighlighted(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemSelected(MenuEvent event)
  {
    // default implementation
  }
}

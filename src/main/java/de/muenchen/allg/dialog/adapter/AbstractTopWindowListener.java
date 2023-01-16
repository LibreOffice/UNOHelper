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

import com.sun.star.awt.XTopWindowListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XTopWindowListener}.
 */
public abstract class AbstractTopWindowListener implements XTopWindowListener
{

  @Override
  public void disposing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowActivated(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowClosed(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowClosing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowDeactivated(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowMinimized(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowNormalized(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowOpened(EventObject arg0)
  {
    // default implementation
  }

}

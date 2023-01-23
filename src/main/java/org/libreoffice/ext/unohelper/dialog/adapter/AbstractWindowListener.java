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
package org.libreoffice.ext.unohelper.dialog.adapter;

import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XWindowListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XWindowListener}.
 */
public abstract class AbstractWindowListener implements XWindowListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void windowHidden(EventObject event)
  {
    // default implementation
  }

  @Override
  public void windowMoved(WindowEvent event)
  {
    // default implementation
  }

  @Override
  public void windowResized(WindowEvent event)
  {
    // default implementation
  }

  @Override
  public void windowShown(EventObject event)
  {
    // default implementation
  }

}

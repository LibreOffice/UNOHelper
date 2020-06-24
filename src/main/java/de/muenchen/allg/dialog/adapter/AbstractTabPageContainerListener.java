/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2020 Landeshauptstadt München
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

import com.sun.star.awt.tab.XTabPageContainerListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XTabPageContainerListener}.
 */
@FunctionalInterface
public interface AbstractTabPageContainerListener extends XTabPageContainerListener
{

  @Override
  public default void disposing(EventObject event)
  {
    // default implementation
  }
}

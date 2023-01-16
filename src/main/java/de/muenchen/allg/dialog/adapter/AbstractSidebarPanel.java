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
package de.muenchen.allg.dialog.adapter;

import com.sun.star.frame.XFrame;
import com.sun.star.lang.XComponent;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.ui.UIElementType;
import com.sun.star.ui.XToolPanel;
import com.sun.star.ui.XUIElement;
import com.sun.star.uno.UnoRuntime;

/**
 * A default implementation for a sidebar panel.
 */
public abstract class AbstractSidebarPanel extends ComponentBase implements XUIElement
{
  /**
   * The resource description.
   */
  private final String resourceUrl;

  /**
   * The panel.
   */
  protected XToolPanel panel;

  /**
   * A default panel. {@link #panel} has to be set manually.
   * 
   * @param resourceUrl
   *          The resource description.
   */
  public AbstractSidebarPanel(String resourceUrl)
  {
    this.resourceUrl = resourceUrl;
  }

  @Override
  public XFrame getFrame()
  {
    return null;
  }

  @Override
  public Object getRealInterface()
  {
    return panel;
  }

  @Override
  public String getResourceURL()
  {
    return resourceUrl;
  }

  @Override
  public short getType()
  {
    return UIElementType.TOOLPANEL;
  }

  @Override
  public void dispose()
  {
    XComponent xPanelComponent = UnoRuntime.queryInterface(XComponent.class, panel);
    xPanelComponent.dispose();
  }
}

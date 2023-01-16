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

import com.sun.star.lang.XServiceInfo;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.ui.XUIElementFactory;
import com.sun.star.uno.XComponentContext;

/**
 * Base implementation for sidebar factories.
 */
public abstract class AbstractSidebarFactory extends WeakBase implements XUIElementFactory, XServiceInfo
{

  /**
   * The context of the factory.
   */
  protected final XComponentContext context;

  @SuppressWarnings("java:S116")
  private final String __serviceName;

  /**
   * Create a new factory.
   * 
   * @param serviceName
   *          The service this factory can create.
   * @param context
   *          The context of the factory.
   */
  public AbstractSidebarFactory(String serviceName, XComponentContext context)
  {
    this.__serviceName = serviceName;
    this.context = context;
  }

  @Override
  public String getImplementationName()
  {
    return this.getClass().getName();
  }

  @Override
  public String[] getSupportedServiceNames()
  {
    return new String[] { __serviceName };
  }

  @Override
  public boolean supportsService(String serviceName)
  {
    for (final String supportedServiceName : getSupportedServiceNames())
    {
      if (supportedServiceName.equals(serviceName))
      {
	return true;
      }
    }
    return false;
  }
}

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
package de.muenchen.allg.util;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.WrappedTargetException;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.afid.UnoProps;

/**
 * Helper for reading and writing the LibreOffice configuration.
 */
public class UnoConfiguration
{
  /**
   * Modify the LibreOffice configuration.
   * 
   * @param nodepath
   *          The path, which is modified.
   * @param props
   *          The properties of the path which are changed.
   * @throws UnoHelperException
   *           The configuration can't be modified.
   */
  public static void setConfiguration(String nodepath, UnoProps props) throws UnoHelperException
  {
    try
    {
      Object cp = UnoComponent.createComponentWithContext(UnoComponent.CSS_CONFIGURATION_CONFIGURATION_PROVIDER);
      Object ca = UnoService.createServiceWithArguments(UnoService.CSS_CONFIGURATION_CONFIGURATION_UPDATE_ACCESS,
          new UnoProps(UnoProperty.NODEPATH, nodepath).getProps(), cp);
      for (PropertyValue prop : props.getProps())
      {
        UnoProperty.setProperty(ca, prop.Name, prop.Value);
      }
      UNO.XChangesBatch(ca).commitChanges();
    } catch (WrappedTargetException e)
    {
      throw new UnoHelperException("Can't modify the configuration", e);
    }
  }

  /**
   * Get a configuration value.
   * 
   * @param nodepath
   *          The path to read from.
   * @param property
   *          The name of the property.
   * @return The value of the property.
   * @throws UnoHelperException
   *           The configuration can't be read.
   */
  public static Object getConfiguration(String nodepath, String property) throws UnoHelperException
  {
    Object cp = UnoComponent.createComponentWithContext(UnoComponent.CSS_CONFIGURATION_CONFIGURATION_PROVIDER);
    Object ca = UnoService.createServiceWithArguments(UnoService.CSS_CONFIGURATION_CONFIGURATION_ACCESS,
        new UnoProps(UnoProperty.NODEPATH, nodepath)
            .getProps(),
        cp);
    return UnoProperty.getProperty(ca, property);
  }

  private UnoConfiguration()
  {
    // nothing to initialize
  }
}

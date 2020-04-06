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

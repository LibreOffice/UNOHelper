package de.muenchen.allg.util;

import com.sun.star.lang.XMultiServiceFactory;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperRuntimeException;

/**
 * Helper for creating UNO services.
 */
@SuppressWarnings("java:S103")
public class UnoService
{

  public static final String CSS_AWT_TREE_MUTABLE_TREE_DATAM_ODEL = "com.sun.star.awt.tree.MutableTreeDataModel";
  public static final String CSS_BEANS_INTROSPECTIONS = "com.sun.star.beans.Introspections";
  public static final String CSS_CONFIGURATION_CONFIGURATION_ACCESS = "com.sun.star.configuration.ConfigurationAccess";
  public static final String CSS_CONFIGURATION_CONFIGURATION_UPDATE_ACCESS = "com.sun.star.configuration.ConfigurationUpdateAccess";
  public static final String CSS_DOCUMENT_SETTINGS = "com.sun.star.document.Settings";
  public static final String CSS_STYLE_CHARACTER_STYLE = "com.sun.star.style.CharacterStyle";
  public static final String CSS_STYLE_PARAGRAPH_STYLE = "com.sun.star.style.ParagraphStyle";
  public static final String CSS_TEXT_BOOKMARK = "com.sun.star.text.Bookmark";
  public static final String CSS_TEXT_FIELD_MASTER_DATABASE = "com.sun.star.text.FieldMaster.Database";
  public static final String CSS_TEXT_FIELD_MASTER_USER = "com.sun.star.text.FieldMaster.User";
  public static final String CSS_TEXT_TEXT_FIELD_ANNOTATION = "com.sun.star.text.TextField.Annotation";
  public static final String CSS_TEXT_TEXT_FIELD_DATA_BASE = "com.sun.star.text.TextField.DataBase";
  public static final String CSS_TEXT_TEXT_FIELD_DATA_BASE_NEXT_SET = "com.sun.star.text.TextField.DataBaseNextSet";
  public static final String CSS_TEXT_TEXT_FIELD_INPUT = "com.sun.star.text.TextField.Input";
  public static final String CSS_TEXT_TEXT_FIELD_INPUT_USER = "com.sun.star.text.TextField.InputUser";
  public static final String CSS_TEXT_TEXT_FIELD_JUMP_EDIT = "com.sun.star.text.TextField.JumpEdit";
  public static final String CSS_TEXT_TEXT_FRAME = "com.sun.star.text.TextFrame";
  public static final String CSS_TEXT_TEXT_SECTION = "com.sun.star.text.TextSection";

  /**
   * Calls {@link #createService(String, XMultiServiceFactory) with {@link UNO#xMSF} .
   *
   * @param serviceName
   *          The name of the service.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createService(String serviceName)
  {
    return createService(serviceName, UNO.xMSF);
  }

  /**
   * Create a service.
   *
   * @param serviceName
   *          The name of the service.
   * @param factory
   *          The service manager.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createService(String serviceName, Object factory)
  {
    return createService(serviceName, UNO.XMultiServiceFactory(factory));
  }

  /**
   * Create a service.
   *
   * @param serviceName
   *          The name of the service.
   * @param factory
   *          The service manager.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createService(String serviceName, XMultiServiceFactory factory)
  {
    try
    {
      return factory.createInstance(serviceName);
    } catch (com.sun.star.uno.Exception e)
    {
      throw new UnoHelperRuntimeException("UNO-Service konnte nicht erstellt werden.", e);
    }
  }

  /**
   * Calls {@link #createServiceWithArguments(String, Object[], XMultiServiceFactory) with
   * {@link UNO#xMSF}.
   *
   * @param serviceName
   *          The name of the service.
   * @param args
   *          The arguments.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createServiceWithArguments(String serviceName, Object[] args)
  {
    return createServiceWithArguments(serviceName, args, UNO.xMSF);
  }

  /**
   * Create a service.
   *
   * @param serviceName
   *          The name of the service.
   * @param args
   *          The arguments.
   * @param factory
   *          The service manager.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createServiceWithArguments(String serviceName, Object[] args, Object factory)
  {
    return createServiceWithArguments(serviceName, args, UNO.XMultiServiceFactory(factory));
  }

  /**
   * Create a service.
   *
   * @param serviceName
   *          The name of the service.
   * @param args
   *          The arguments.
   * @param factory
   *          The service manager.
   * @return The service.
   * @throws UnoHelperRuntimeException
   *           Service can't be created.
   */
  public static Object createServiceWithArguments(String serviceName, Object[] args, XMultiServiceFactory factory)
  {
    try
    {
      return factory.createInstanceWithArguments(serviceName, args);
    } catch (com.sun.star.uno.Exception e)
    {
      throw new UnoHelperRuntimeException("UNO-Service konnte nicht erstellt werden.", e);
    }
  }

  /**
   * Check if an objects is a service.
   *
   * @param object
   *          The object.
   * @param serviceName
   *          The name of the service.
   * @return True if the object is such a service, false otherwise.
   */
  public static boolean supportsService(Object object, String serviceName)
  {
    if (UNO.XServiceInfo(object) != null)
    {
      return UNO.XServiceInfo(object).supportsService(serviceName);
    }
    return false;
  }

  private UnoService()
  {
    // nothing to do
  }
}

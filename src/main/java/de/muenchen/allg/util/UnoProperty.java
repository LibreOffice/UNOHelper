package de.muenchen.allg.util;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertyState;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;

/**
 * Helper for accessing and modifying properties of UNO objects.
 */
@SuppressWarnings("java:S1845")
public final class UnoProperty
{

  public static final String ACTION_COMMAND = "ActionCommand";
  public static final String ACTIVE_CONNECTION = "ActiveConnection";
  public static final String ANCHOR_TYPE = "AnchorType";
  public static final String AS_TEMPLATE = "AsTemplate";
  public static final String AUTHOR = "Author";
  public static final String BACK_TRANSPARENT = "BackTransparent";
  public static final String BACKGROUND_COLOR = "BackgroundColor";
  public static final String BOOKMARK = "Bookmark";
  public static final String BORDER = "Border";
  public static final String BORDER_COLOR = "BorderColor";
  public static final String BORDER_DISTANCE = "BorderDistance";
  public static final String BOTTOM_BORDER = "BottomBorder";
  public static final String BREAK_TYPE = "BreakType";
  public static final String CHAR_BACK_COLOR = "CharBackColor";
  public static final String CHAR_COLOR = "CharColor";
  public static final String CHAR_FONT_NAME = "CharFontName";
  public static final String CHAR_HEIGHT = "CharHeight";
  public static final String CHAR_HIDDEN = "CharHidden";
  public static final String CHAR_SET = "CharSet";
  public static final String CHAR_STYLE_NAME = "CharStyleName";
  public static final String CHAR_UNDERLINE = "CharUnderline";
  public static final String CHAR_WEIGHT = "CharWeight";
  public static final String COLLATE = "Collate";
  public static final String COMMAND = "Command";
  public static final String COMMAND_TYPE = "CommandType";
  public static final String COMMAND_URL = "CommandURL";
  public static final String CONDITION = "Condition";
  public static final String CONTENT = "Content";
  public static final String COPY_COUNT = "CopyCount";
  public static final String CURRENT_PRESENTAITON = "CurrentPresentation";
  public static final String DATA_BASE_NAME = "DataBaseName";
  public static final String DATA_COLUMN_NAME = "DataColumnName";
  public static final String DATA_COMMAND_TYPE = "DataCommandType";
  public static final String DATA_MODEL = "DataModel";
  public static final String DATA_SOURCE_NAME = "DataSourceName";
  public static final String DATA_TABLE_NAME = "DataTableName";
  public static final String DECIMAL_ACCURACY = "DecimalAccuracy";
  public static final String DECIMAL_DELIMITER = "DecimalDelimiter";
  public static final String DEFAULT_CONTEXT = "DefaultContext";
  public static final String DEFAULT_CONTROL = "DefaultControl";
  public static final String DOCUMENT_URL = "DocumentURL";
  public static final String DROPDOWN = "Dropdown";
  public static final String EMPTY_PAGES = "EmptyPages";
  public static final String ENABLE_VISIBLE = "EnableVisible";
  public static final String ENABLED = "Enabled";
  public static final String ESCAPE_PROCESSING = "EscapeProcessing";
  public static final String EXTENSION = "Extension";
  public static final String FIELD_DELIMITER = "FieldDelimiter";
  public static final String FILE_NAME_FROM_COLUMN = "FileNameFromColumn";
  public static final String FILE_NAME_PREFIX = "FileNamePrefix";
  public static final String FILTER_NAME = "FilterName";
  public static final String FIXED_LENGTH = "FixedLength";
  public static final String FOLLOW_STYLE = "FollowStyle";
  public static final String FONT_DESCRIPTOR = "FontDescriptor";
  public static final String FRAME_IS_AUTOMATIC_HEIGHT = "FrameIsAutomaticHeight";
  public static final String HEADER_LINE = "HeaderLine";
  public static final String HEIGHT = "Height";
  public static final String HELP_TEXT = "HelpText";
  public static final String HIDDEN = "Hidden";
  public static final String HIDE_INACTIVE_SELECTION = "HideInactiveSelection";
  public static final String HINT = "Hint";
  public static final String HORI_ORIENT = "HoriOrient";
  public static final String HORI_ORIENT_POSITION = "HoriOrientPosition";
  public static final String HORI_ORIENT_RELATION = "HoriOrientRelation";
  public static final String IMAGE_URL = "ImageURL";
  public static final String INFO = "Info";
  public static final String INPUT_STREAM = "InputStream";
  public static final String IS_COLLAPSED = "IsCollapsed";
  public static final String IS_START = "IsStart";
  public static final String IS_VISIBLE = "IsVisible";
  public static final String ITEM_DESCRIPTOR_CONTAINER = "ItemDescriptorContainer";
  public static final String ITEMS = "Items";
  public static final String LABEL = "Label";
  public static final String LAYOUT_MANAGER = "LayoutManager";
  public static final String LEFT_BORDER = "LeftBorder";
  public static final String LOAD_CELL_STYLES = "LoadCellStyles";
  public static final String LOAD_FRAME_STYLES = "LoadFrameStyles";
  public static final String LOAD_NUMBERING_STYLES = "LoadNumberingStyles";
  public static final String LOAD_PAGE_STYLES = "LoadPageStyles";
  public static final String LOAD_PRINTER = "LoadPrinter";
  public static final String LOAD_TEXT_STYLES = "LoadTextStyles";
  public static final String MACRO_EXECUTION_MODE = "MacroExecutionMode";
  public static final String MAX_VALUE = "MaxValue";
  public static final String MIN_VALUE = "MinValue";
  public static final String MULTILINE = "MultiLine";
  public static final String NAME = "Name";
  public static final String NODEPATH = "nodepath";
  public static final String OO_SETUP_EXTENSION = "ooSetupExtension";
  public static final String OO_SETUP_FACTORY_WINDOW_ATTRIBUTES = "ooSetupFactoryWindowAttributes";
  public static final String OO_SETUP_VERSION_ABOUT_BOX = "ooSetupVersionAboutBox";
  public static final String OUTPUT_TYPE = "OutputType";
  public static final String OUTPUT_URL = "OutputURL";
  public static final String OVERWRITE_STYLES = "OverwriteStyles";
  public static final String PAGE_COUNT = "PageCount";
  public static final String PAGES = "Pages";
  public static final String PARA_FIRST_LINE_INDENT = "ParaFirstLineIndent";
  public static final String PARA_STYLE_NAME = "ParaStyleName";
  public static final String PARA_TOP_MARGIN = "ParaTopMargin";
  public static final String POSITION_X = "PositionX";
  public static final String POSITION_Y = "PositionY";
  public static final String PRINT_EMPTY_PAGES = "PrintEmptypages";
  public static final String READ_ONLY = "ReadOnly";
  public static final String RECORD_CHANGES = "RecordChanges";
  public static final String RIGHT_BORDER = "RightBorder";
  public static final String ROOT_DISPLAYED = "RootDisplayed";
  public static final String ROW_HEIGHT = "RowHeight";
  public static final String SAVE_AS_SINGLE_FILE = "SaveAsSingleFile";
  public static final String SAVE_FILTER = "SaveFilter";
  public static final String SELECTED_ITEM = "SelectedItem";
  public static final String SHOW_THOUSANDS_SEPERATOR = "ShowThousandsSeperator";
  public static final String SINGLE_PRINT_JOBS = "SinglePrintJobs";
  public static final String SPIN = "Spin";
  public static final String STATE = "State";
  public static final String STRING_DELIMITER = "StringDelimiter";
  public static final String TABINDEX = "TabIndex";
  public static final String TEXT = "Text";
  public static final String TEXT_COLOR = "TextColor";
  public static final String TEXT_FIELD = "TextField";
  public static final String TEXT_PROTION_TYPE = "TextPortionType";
  public static final String TEXT_WRAP = "TextWrap";
  public static final String THOUSAND_DELIMITER = "ThousandDelimiter";
  public static final String TITLE = "Title";
  public static final String TOGGLE = "Toggle";
  public static final String TOP_BORDER = "TopBorder";
  public static final String TYPE = "Type";
  public static final String URI = "URI";
  public static final String URL = "URL";
  public static final String VALUE = "Value";
  public static final String VERT_ORIENT = "VertOrient";
  public static final String VERT_ORIENT_POSITION = "VertOrientPosition";
  public static final String VERT_ORIENT_RELATION = "VertOrientRelation";
  public static final String WAIT = "Wait";
  public static final String WIDTH = "Width";
  public static final String WORK = "Work";
  public static final String ZOOM_TYPE = "ZoomType";
  public static final String ZOOM_VALUE = "ZoomValue";

  /**
   * Get the value of a property from the object.
   *
   * @param object
   *          The object which has the property.
   * @param propName
   *          The name of the requested property.
   * @return The value of the property.
   * @throws UnoHelperException
   *           If the object doesn't support {@link XPropertySet} or no property with this name
   *           exists.
   */
  public static Object getProperty(Object object, String propName) throws UnoHelperException
  {
    try
    {
      XPropertySet props = UNO.XPropertySet(object);
      if (props == null)
      {
        throw new UnoHelperException("Object doesn't support XPropertySet.");
      }
      return props.getPropertyValue(propName);
    } catch (UnknownPropertyException | WrappedTargetException e)
    {
      throw new UnoHelperException("Property can't be read.", e);
    }
  }

  /**
   * Get the value of a property from the given properties.
   *
   * @param propertyValues
   *          The some properties
   * @param propertyName
   *          The name of the requested property.
   * @return The value of the property, {@code null} if there's no property with the requested name.
   */
  public static String getPropertyByPropertyValues(PropertyValue[] propertyValues, String propertyName)
  {
    String value = null;

    for (PropertyValue property : propertyValues)
    {
      if (property.Name.equals(propertyName))
      {
        value = (String) property.Value;
        break;
      }
    }

    return value;
  }

  /**
   * Set a property of the object to the given value.
   *
   * @param object
   *          The object which has the property.
   * @param propName
   *          The name of the property.
   * @param propVal
   *          The new value of the property.
   * @return The value after the property was modified.
   * @throws UnoHelperException
   *           If the object doesn't support {@link XPropertySet} or no property with this name
   *           exists or the property can't be set.
   */
  public static Object setProperty(Object object, String propName, Object propVal) throws UnoHelperException
  {
    XPropertySet props = UNO.XPropertySet(object);
    if (props == null)
    {
      throw new UnoHelperException("Object doesn't support XPropertySet.");
    }
    try
    {
      props.setPropertyValue(propName, propVal);
      return props.getPropertyValue(propName);
    } catch (UnknownPropertyException | IllegalArgumentException | PropertyVetoException | WrappedTargetException e)
    {
      throw new UnoHelperException("Property can't be written.", e);
    }
  }

  /**
   * Set a property of the object to its default value.
   *
   * @param object
   *          The object which has the property.
   * @param propName
   *          The name of the property.
   * @return The value after the property was modified.
   * @throws UnoHelperException
   *           If the object doesn't support {@link XPropertyState} or no property with this name
   *           exists or the property can't be set.
   */
  public static Object setPropertyToDefault(Object object, String propName) throws UnoHelperException
  {
    try
    {
      XPropertyState props = UNO.XPropertyState(object);
      if (props == null)
      {
        throw new UnoHelperException("Object doesn't support XPropertyState.");
      }
      props.setPropertyToDefault(propName);
      return getProperty(props, propName);
    } catch (UnknownPropertyException e)
    {
      throw new UnoHelperException("Property can't be written.", e);
    }
  }

  private UnoProperty()
  {
    // nothing to do
  }
}

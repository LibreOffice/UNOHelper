package de.muenchen.allg.afid;

import java.util.Arrays;
import java.util.stream.Collectors;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;

/**
 * Wrapper for {@link PropertyValue}-arrays.
 */
public class UnoProps
{

  /**
   * The properties.
   */
  private PropertyValue[] props;

  /**
   * Wraps an empty {@link PropertyValue}-array.
   */
  public UnoProps()
  {
    props = new PropertyValue[] {};
  }

  /**
   * Wraps the given values as a {@link PropertyValue}-array.
   *
   * @param values
   *          Some {@link PropertyValue}s.
   */
  public UnoProps(PropertyValue... values)
  {
    props = values;
  }

  /**
   * Create a new {@link PropertyValue}-array and add a property.
   *
   * @param name
   *          The name of the property.
   * @param value
   *          The value of the property.
   */
  public UnoProps(String name, Object value)
  {
    this();
    setPropertyValue(name, value);
  }

  /**
   * Create a new {@link PropertyValue}-array and add a property.
   *
   * @param name1
   *          The name of the first property.
   * @param value1
   *          The value of the first property.
   * @param name2
   *          The name of the second property.
   * @param value2
   *          The value of the second property.
   */
  public UnoProps(String name1, Object value1, String name2, Object value2)
  {
    this(name1, value1);
    setPropertyValue(name2, value2);
  }

  /**
   * Get the properties as sorted array.
   * 
   * @return The {@link PropertyValue}-array.
   */
  public PropertyValue[] getProps()
  {
    Arrays.sort(props, (x, y) -> x.Name.compareTo(y.Name));
    return props;
  }

  /**
   * Add or overwrite a property.
   * 
   * @param name
   *          The name of the property.
   * @param value
   *          The value of the property.
   * @return This UnoProps (fluent API)
   */
  public UnoProps setPropertyValue(String name, Object value)
  {
    for (PropertyValue prop : props)
    {
      if (prop != null && prop.Name.equals(name))
      {
        prop.Value = value;
        return this;
      }
    }

    // create new entry
    props = Arrays.copyOf(props, props.length + 1);
    PropertyValue prop = new PropertyValue();
    prop.Name = name;
    prop.Value = value;
    props[props.length - 1] = prop;
    return this;
  }

  /**
   * Get the value of a property.
   * 
   * @param name
   *          The name of the property.
   * @return The value of the property.
   * @throws UnknownPropertyException
   *           There's no property with this name.
   */
  public Object getPropertyValue(String name) throws UnknownPropertyException
  {
    for (int i = 0; i < props.length; i++)
    {
      if (props[i].Name.equals(name))
        return props[i].Value;
    }
    throw new UnknownPropertyException(name);
  }

  /**
   * Get the value of a property as string.
   * 
   * @param name
   *          The name of the property.
   * @return The value of the property as string ({@link Object#toString()}).
   * @throws UnknownPropertyException
   *           There's no property with this name.
   */
  public String getPropertyValueAsString(String name) throws UnknownPropertyException
  {
    return getPropertyValue(name).toString();
  }

  @Override
  public String toString()
  {
    return "UnoProps[ "
        + Arrays.stream(props).map(p -> p.Name + "=>\"" + p.Value + "\"").collect(Collectors.joining(", ")) + " ]";
  }
}

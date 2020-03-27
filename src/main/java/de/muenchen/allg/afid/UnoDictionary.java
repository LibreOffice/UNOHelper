package de.muenchen.allg.afid;

import java.util.AbstractMap;
import java.util.HashSet;
import java.util.Set;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;

/**
 * Wrapper for {@link XNameAccess} so that it can be accessed like a dictionary.
 * 
 * @param <V>
 *          Type of objects in the dictionary.
 */
public class UnoDictionary<V> extends AbstractMap<String, V>
{
  private XNameAccess access;

  private Set<Entry<String, V>> entries;

  /**
   * Create a dictionary for the object.
   *
   * @param o
   *          An object which implements {@link XNameAccess}.
   */
  @SuppressWarnings("unchecked")
  public UnoDictionary(Object o)
  {
    access = UNO.XNameAccess(o);
    entries = new HashSet<>();
    for (String name : access.getElementNames())
    {
      entries.add(new Entry<String, V>()
      {

        @Override
        public String getKey()
        {
          return name;
        }

        @Override
        public V getValue()
        {
          try
          {
            return (V) access.getByName(name);
          } catch (NoSuchElementException | WrappedTargetException e)
          {
            return null;
          }
        }

        @Override
        public V setValue(V arg0)
        {
          throw new UnsupportedOperationException();
        }
      });
    }
  }

  /**
   * @see #containsKey(Object)
   * @deprecated
   */
  @Deprecated(since = "3.0.0", forRemoval = true)
  public boolean hasKey(String key)
  {
    return containsKey(key);
  }

  @Override
  public Set<Entry<String, V>> entrySet()
  {
    return entries;
  }

  @Override
  public boolean equals(Object o)
  {
    boolean equal = super.equals(o);
    if (equal && o instanceof UnoDictionary)
    {
      UnoDictionary<?> other = (UnoDictionary<?>) o;
      return UnoRuntime.areSame(access, other.access);
    }
    return false;
  }

  @Override
  public int hashCode()
  {
    return super.hashCode() + access.hashCode();
  }

  /**
   * Get the type of objects.
   *
   * @return The type of objects in the dictionary.
   */
  public Type getElementType()
  {
    return access.getElementType();
  }
}

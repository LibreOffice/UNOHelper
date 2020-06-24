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
  /**
   * The dictionary.
   */
  private XNameAccess access;

  /**
   * Contains all entries. Initialized in {@link #entrySet()}.
   */
  private Set<Entry<String, V>> entries;

  /**
   * The type of objects in the enumeration.
   */
  private Class<V> c;

  /**
   * Create a dictionary for the object.
   *
   * @param o
   *          An object which implements {@link XNameAccess}.
   * @param c
   *          The type of objects in the enumeration.
   */
  private UnoDictionary(XNameAccess access, Class<V> c)
  {
    this.access = access;
    this.c = c;
    entries = new HashSet<>();
  }

  /**
   * Create an UnoDictionary for the provided mapping.
   *
   * @param <T>
   *          The type of the objects in the mapping.
   * @param access
   *          The mapping.
   * @param c
   *          The class of the type.
   * @return An UnoDictionary.
   */
  public static <T> UnoDictionary<T> create(XNameAccess access, Class<T> c)
  {
    return new UnoDictionary<>(access, c);
  }

  /**
   * Create an UnoDictionary for the provided object. The object must implement {@link XNameAccess}.
   *
   * @param <T>
   *          The type of the objects in the enumeration.
   * @param o
   *          The mapping.
   * @param c
   *          The class of the type.
   * @return An UnoDictionary or null if object doesn't implement {@link XNameAccess}.
   */
  public static <T> UnoDictionary<T> create(Object o, Class<T> c)
  {
    XNameAccess access = UNO.XNameAccess(o);
    if (access != null)
    {
      return new UnoDictionary<>(access, c);
    }

    return null;
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
  public boolean containsKey(Object key)
  {
    return access.hasByName(key.toString());
  }

  @Override
  @SuppressWarnings("unchecked")
  public V get(Object key)
  {
    String name = key.toString();
    try
    {
      if (c != Object.class)
      {
        return UnoRuntime.queryInterface(c, access.getByName(name));
      } else
      {
        return (V) access.getByName(name);
      }
    } catch (NoSuchElementException | WrappedTargetException e)
    {
      return null;
    }
  }

  @Override
  public Set<String> keySet()
  {
    return Set.of(access.getElementNames());
  }

  @Override
  @SuppressWarnings("unchecked")
  public Set<Entry<String, V>> entrySet()
  {
    if (entries.isEmpty())
    {
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
              if (c != Object.class)
              {
                return UnoRuntime.queryInterface(c, access.getByName(name));
              } else
              {
                return (V) access.getByName(name);
              }
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

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
package org.libreoffice.unohelper.common;

import java.util.AbstractList;

import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

/**
 * Wrapper for {@link XIndexAccess}.
 *
 * @param <T>
 *          type of objects in the list.
 */
public class UnoList<T> extends AbstractList<T>
{
  /**
   * The access to the objects.
   */
  private XIndexAccess access;

  /**
   * The type of objects.
   */
  private Class<T> c;

  /**
   * Create new UnoList.
   *
   * @param access
   *          An object which implements {@link XIndexAccess}.
   * @param c
   *          The type of the objects.
   */
  private UnoList(XIndexAccess access, Class<T> c)
  {
    this.access = access;
    this.c = c;
  }

  /**
   * Create an UnoList for the provided mapping.
   *
   * @param <T>
   *          The type of the objects in the mapping.
   * @param access
   *          The mapping.
   * @param c
   *          The type of the objects.
   * @return An UnoList.
   */
  public static <T> UnoList<T> create(XIndexAccess access, Class<T> c)
  {
    return new UnoList<>(access, c);
  }

  /**
   * Create an UnoList for the provided object. The object must implement {@link XNameAccess}.
   *
   * @param <T>
   *          The type of the objects in the enumeration.
   * @param o
   *          The mapping.
   * @param c
   *          The type of the objects.
   * @return An UnoList or null if object doesn't implement {@link XNameAccess}.
   */
  public static <T> UnoList<T> create(Object o, Class<T> c)
  {
    XIndexAccess access = UNO.XIndexAccess(o);
    if (access != null)
    {
      return new UnoList<>(access, c);
    }

    return null;
  }

  public XIndexAccess getAccess()
  {
    return access;
  }

  @SuppressWarnings("unchecked")
  @Override
  public T get(int index)
  {
    try
    {
      if (c != Object.class)
      {
        return UnoRuntime.queryInterface(c, access.getByIndex(index));
      } else
      {
        return (T) access.getByIndex(index);
      }
    } catch (IndexOutOfBoundsException | WrappedTargetException e)
    {
      return null;
    }
  }

  @Override
  public int size()
  {
    return access.getCount();
  }

  @Override
  public boolean equals(Object o)
  {
    boolean equal = super.equals(o);
    if (equal && o instanceof UnoList)
    {
      UnoList<?> other = (UnoList<?>) o;
      return UnoRuntime.areSame(access, other.access);
    }
    return false;
  }

  @Override
  public int hashCode()
  {
    return super.hashCode() + access.hashCode();
  }

}

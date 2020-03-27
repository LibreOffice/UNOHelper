package de.muenchen.allg.afid;

import java.util.AbstractList;

import com.sun.star.container.XIndexAccess;
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
   * Create new UnoList.
   *
   * @param o
   *          An object which implements {@link XIndexAccess}.
   */
  public UnoList(Object o)
  {
    access = UNO.XIndexAccess(o);
  }

  @SuppressWarnings("unchecked")
  @Override
  public T get(int index)
  {
    try
    {
      return (T) access.getByIndex(index);
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

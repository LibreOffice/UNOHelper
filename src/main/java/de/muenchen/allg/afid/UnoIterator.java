package de.muenchen.allg.afid;

import java.util.Iterator;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.uno.UnoRuntime;

/**
 * Makes an enumeration iterable.
 *
 * @param <T>
 *          type of the objects in the enumeration.
 */
public class UnoIterator<T> implements Iterator<T>
{
  /**
   * The enumeration.
   */
  private XEnumeration enu;

  /**
   * The type of objects in the enumeration.
   */
  private Class<T> c;

  /**
   * Create an UnoIterator.
   *
   * @param enuAccess
   *          The enumeration.
   * @param c
   *          The type of objects in the enumeration.
   */
  public UnoIterator(XEnumerationAccess enuAccess, Class<T> c)
  {
    this.c = c;
    enu = enuAccess.createEnumeration();
  }

  @Override
  public boolean hasNext()
  {
    return enu.hasMoreElements();
  }

  @SuppressWarnings("unchecked")
  @Override
  public T next()
  {
    try
    {
      if (c != Object.class)
      {
        return UnoRuntime.queryInterface(c, enu.nextElement());
      } else
      {
        return (T) enu.nextElement();
      }
    } catch (NoSuchElementException e)
    {
      throw new java.util.NoSuchElementException();
    } catch (Exception e)
    {
      return null;
    }
  }

  /**
   * Unsupported operation.
   */
  @Override
  public void remove()
  {
    throw new UnsupportedOperationException("Remove is not allowed");
  }

}

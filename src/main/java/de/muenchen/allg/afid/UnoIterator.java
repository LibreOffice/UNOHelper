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
   * @param enu
   *          The enumeration.
   * @param c
   *          The type of objects in the enumeration.
   */
  private UnoIterator(XEnumeration enu, Class<T> c)
  {
    this.c = c;
    this.enu = enu;
  }

  /**
   * Create an UnoIterator for the provided enumeration.
   *
   * @param <T>
   *          The type of the objects in the enumeration.
   * @param enuAccess
   *          The enumeration.
   * @param c
   *          The class of the type.
   * @return An UnoIterator.
   */
  public static <T> UnoIterator<T> create(XEnumerationAccess enuAccess, Class<T> c)
  {
    return new UnoIterator<>(enuAccess.createEnumeration(), c);
  }

  /**
   * Create an UnoIterator for the provided enumeration.
   *
   * @param <T>
   *          The type of the objects in the enumeration.
   * @param enu
   *          The enumeration.
   * @param c
   *          The class of the type.
   * @return An UnoIterator.
   */
  public static <T> UnoIterator<T> create(XEnumeration enu, Class<T> c)
  {
    return new UnoIterator<>(enu, c);
  }

  /**
   * Create an UnoIterator for the provided object. The object must implement
   * {@link XEnumerationAccess}.
   *
   * @param <T>
   *          The type of the objects in the enumeration.
   * @param o
   *          The enumeration.
   * @param c
   *          The class of the type.
   * @return An UnoIterator or null if object doesn't implement {@link XEnumerationAccess}.
   */
  public static <T> UnoIterator<T> create(Object o, Class<T> c)
  {
    XEnumerationAccess enuAccess = UNO.XEnumerationAccess(o);
    if (enuAccess != null)
    {
      return new UnoIterator<>(enuAccess.createEnumeration(), c);
    }

    return null;
  }

  @Override
  public boolean hasNext()
  {
    return enu != null && enu.hasMoreElements();
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

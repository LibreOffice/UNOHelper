package de.muenchen.allg.afid;

import java.util.Iterator;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.uno.UnoRuntime;

public class UnoIterator<T> implements Iterator<T>
{
  private XEnumeration enu;
  private Class<T> c;

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

  @Override
  public T next()
  {
    try
    {
      if (c != Object.class)
      {
        return UnoRuntime.queryInterface(c, enu.nextElement());
      }
      else
      {
        return (T)enu.nextElement();
      }
    } catch (NoSuchElementException e)
    {
      throw new java.util.NoSuchElementException();
    } catch (Exception e)
    {
      return null;
    }
  }

}

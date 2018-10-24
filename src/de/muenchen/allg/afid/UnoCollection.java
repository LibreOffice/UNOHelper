package de.muenchen.allg.afid;

import java.util.Iterator;

import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;

/**
 * Wrapperklasse für {@link XEnumeration}. Kann mit for-each-loop iteriert werden.
 * 
 * @author snoopy
 *
 * @param <T>
 */
public class UnoCollection<T> implements Iterable<T>
{
  private XEnumerationAccess enuAccess;
  private Class<T> c;

  protected UnoCollection(XEnumerationAccess enuAccess, Class<T> c)
  {
    this.enuAccess = enuAccess;
    this.c = c;

  }

  /**
   * Erzeugt eine UnoCollection aus einem Uno-Object, dass XEnumerationAccess implementiert.
   * 
   * @param enuAccess
   * @param c darf nur ein Interface aus der UNO-Api oder Object sein.
   * @return
   */
  public static <T> UnoCollection<T> getCollection(XEnumerationAccess enuAccess, Class<T> c)
  {
    return new UnoCollection<T>(enuAccess, c);
  }
  
  public static <T> UnoCollection<T> getCollection(Object o, Class<T> c)
  {
    XEnumerationAccess enuAccess = UNO.XEnumerationAccess(o);
    if (enuAccess != null)
    {    
    	return new UnoCollection<T>(enuAccess, c);
    }
    
    return null;
  }

  @Override
  public Iterator<T> iterator()
  {
    return new UnoIterator<T>(enuAccess, c);
  }
}

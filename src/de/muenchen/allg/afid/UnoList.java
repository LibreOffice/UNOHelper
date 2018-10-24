package de.muenchen.allg.afid;

import java.util.AbstractList;

import com.sun.star.container.XIndexAccess;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;

public class UnoList<T> extends AbstractList<T>
{
	private XIndexAccess access;

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

}

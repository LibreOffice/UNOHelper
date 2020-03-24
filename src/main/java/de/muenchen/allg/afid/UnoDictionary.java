package de.muenchen.allg.afid;

import java.util.Arrays;
import java.util.Collections;
import java.util.Dictionary;
import java.util.Enumeration;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.Type;

/**
 * Wrappt ein UNO-XNameAcess als JAVA-Dictionary.
 * 
 * @param <V>
 *            Typ des Values im Dictionary
 */
public class UnoDictionary<V> extends Dictionary<String, V>
{
	private XNameAccess access;

	public UnoDictionary(Object o)
	{
		access = UNO.XNameAccess(o);
	}

	/**
	 * Wird nicht unterstützt.
	 */
	@Override
	public Enumeration<V> elements()
	{
		throw new UnsupportedOperationException();
	}

	@SuppressWarnings("unchecked")
	@Override
	public V get(Object key)
	{
		try
		{
			return (V) access.getByName(key.toString());
		} catch (NoSuchElementException | WrappedTargetException e)
		{
			return null;
		}
	}

	@Override
	public boolean isEmpty()
	{
		return !access.hasElements();
	}

	@Override
	public Enumeration<String> keys()
	{
		return Collections.enumeration(Arrays.asList(access.getElementNames()));
	}

	/**
	 * Wird nicht unterstützt.
	 */
	@Override
	public V put(String arg0, V arg1)
	{
		throw new UnsupportedOperationException();
	}

	/**
	 * Wird nicht unterstützt.
	 */
	@Override
	public V remove(Object arg0)
	{
		throw new UnsupportedOperationException();
	}

	/**
	 * Wird nicht unterstützt.
	 */
	@Override
	public int size()
	{
		throw new UnsupportedOperationException();
	}

	public boolean hasKey(Object key)
	{
		return access.hasByName(key.toString());
	}

	public Type getElementType()
	{
		return access.getElementType();
	}
}

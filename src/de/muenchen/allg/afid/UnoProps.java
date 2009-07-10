/*
 * Dateiname: UnoProps.java
 * Projekt  : UNOHelper
 * Funktion : Erleichtert den Umgang mit PropertyValue-Arrays, die in der
 *            UNO-Schnittstelle häufig als Übergabeparameter verwendet werden.
 * 
 * Copyright (c) 2008 Landeshauptstadt München
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL),
 * version 1.0.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * European Union Public Licence for more details.
 *
 * You should have received a copy of the European Union Public Licence
 * along with this program. If not, see
 * http://ec.europa.eu/idabc/en/document/7330
 *
 * Änderungshistorie:
 * Datum      | Wer | Änderungsgrund
 * -------------------------------------------------------------------
 * 16.11.2005 | LUT | Erstellung
 * -------------------------------------------------------------------
 *
 * @author Christoph Lutz (D-III-ITD 5.1)
 * @version 1.0
 * 
 */
package de.muenchen.allg.afid;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;

/**
 * Diese Klasse erleichtert den Umgang mit PropertyValue-Arrays, die in der
 * UNO-Schnittstelle häufig als Übergabeparameter verwendet werden.
 * 
 * @author Christoph Lutz (D-III-ITD 5.1)
 */
public class UnoProps {

	private PropertyValue[] props;

	/**
	 * Der Konstruktor erzeugt und kapselt ein leeres PropertyValue-Array.
	 */
	public UnoProps() {
		props = new PropertyValue[] {};
	}

	/**
	 * Der Konstruktor erzeugt und kapselt ein neues PropertyValue-Array und
	 * belegt es sofort mit dem überbebenen Property.
	 * 
	 * @param name
	 * @param value
	 */
	public UnoProps(String name, Object value) {
		props = new PropertyValue[] {};
		setPropertyValue(name, value);
	}

	/**
	 * Der Konstruktor erzeugt und kapselt ein neues PropertyValue-Array belegt
	 * es sofort mit den zwei übergebenen Properties.
	 * 
	 * @param name1
	 * @param value1
	 * @param name2
	 * @param value2
	 */
	public UnoProps(String name1, Object value1, String name2, Object value2) {
		props = new PropertyValue[] {};
		setPropertyValue(name1, value1);
		setPropertyValue(name2, value2);
	}

	/**
	 * Der Konstruktor kapselt ein bestehendes PropertyValue-Array, das in Form
	 * eines Objec-Arrays übergeben wird. Alle Elemente, die nicht vom Typ
	 * PropertyValue sind, werden ohne Fehlermeldung ignoriert.
	 * 
	 * @param props
	 */
  public UnoProps(Object[] props)
  {
    if (props == null)
    {
      this.props = new PropertyValue[] {};
      return;
    }

    // Zähle Anzahl der PropertyValue-Einträge.
    int count = 0;
    for (int i = 0; i < props.length; i++)
    {
      if (props[i] instanceof PropertyValue) count++;
    }

    // Erzeuge neues PropertyValue[]:
    this.props = new PropertyValue[count];
    int x = 0;
    for (int i = 0; i < props.length; i++)
    {
      if (props[i] instanceof PropertyValue)
      {
        PropertyValue pv = (PropertyValue) props[i];
        this.props[x++] = new PropertyValue(pv.Name, pv.Handle, pv.Value, pv.State);
      }
    }
  }

	/**
	 * Liefert das zugehörige PropertyValue[] zurück.
	 * 
	 * @return
	 */
	public PropertyValue[] getProps() {
		return props;
	}

	/**
	 * Setzt den Wert value der Property mit dem Namen name. Ist das Property
	 * bereits definiert, so wird der bestehende Wert überschrieben. Ist das
	 * Property noch nicht definiert, so wird ein neuer Eintrag im
	 * PropertyValue-Array erzeugt.
	 * 
	 * @param name
	 *            der Name des Properties.
	 * @param value
	 *            der neue Wert des Properties.
	 * @return Liefert sich selbst zurück, um mehrere Properties in einer Kette
	 *         setzen zu können.
	 */
	public UnoProps setPropertyValue(String name, Object value) {
		// Suche nach dem Property erstelle gleich eine Kopie für später:
		PropertyValue[] newProps = new PropertyValue[props.length + 1];
		PropertyValue prop = null;
		for (int i = 0; i < props.length; i++) {
			if (props[i] != null && props[i].Name.equals(name))
				prop = props[i];
			newProps[i] = props[i];
		}

		if (prop != null) {
			// überschreibe den alten Wert:
			prop.Value = value;
		} else {
			// erzeuge einen neuen Eintrag:
			prop = new PropertyValue();
			prop.Name = name;
			prop.Value = value;
			newProps[props.length] = prop;
			props = newProps;
		}
		return this;
	}

	/**
	 * Liefert den Wert eines Properties mit dem Namen name zurück.
	 * 
	 * @param name
	 * @return den Wert des Properties.
	 * @throws UnknownPropertyException
	 */
	public Object getPropertyValue(String name) throws UnknownPropertyException {
		for (int i = 0; i < props.length; i++) {
			if (props[i].Name.equals(name))
				return props[i].Value;
		}
		throw new UnknownPropertyException(name);
	}

	/**
	 * Liefert ein UnoService-Objekt zurück, das den den Wert des Properties mit
	 * dem Namen name kapselt.
	 * 
	 * @param name
	 * @return Liefert ein UnoService-Objekt zurück, das den den Wert des
	 *         Properties mit dem Namen name kapselt.
	 * @throws UnknownPropertyException
	 */
	public UnoService getPropertyValueAsUnoService(String name)
			throws UnknownPropertyException {
		return new UnoService(getPropertyValue(name));
	}

	/**
	 * Liefert ein String mit dem Wert des Properties mit dem Namen name zurück.
	 * 
	 * @param name
	 * @return Liefert ein String mit dem Wert des Properties mit dem Namen name
	 *         zurück.
	 * @throws UnknownPropertyException
	 */
	public String getPropertyValueAsString(String name)
			throws UnknownPropertyException {
		return getPropertyValue(name).toString();
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see java.lang.Object#toString()
	 */
	public String toString() {
		String str = "UnoProps[ ";
		for (int i = 0; i < props.length; i++) {
			if (i != 0)
				str += ", ";
			str += props[i].Name + "=>\"" + props[i].Value + "\"";
		}
		str += " ]";
		return str;
	}
}

/*
 * Ein einfacher Wrapper für uno-services.
 * Copyright (C) 2005 Christoph Lutz
 *
 * Diese Software wurde ursprünglich privat von Christoph Lutz entwickelt
 * und unter einer LGPL-Lizenz freigegeben. Quelle: 
 * http://codesnippets.services.openoffice.org/Office/Office.SimpleWrapperForUno_services.snip
 * 
 * Der Autor berechtigt die Landeshauptstadt München, darüber hinaus, die hier 
 * vorliegende Kopie der Software im vollen Umfang uneingeschränkt ganz oder 
 * in Auszügen zu benutzen, zu verändern und weiterzugeben.
 * 
 * 
* Änderungshistorie der Landeshauptstadt München:
* Alle Änderungen (c) Landeshauptstadt München, alle Rechte vorbehalten
* 
* Nr. |  Datum     |   Autor   | Änderungsgrund
* -------------------------------------------------------------------
* 001 | 26.04.2005 |    BNK    | getSimpleName() durch getName() 
*                              | ersetzt wg. Java 1.4 Kompatibilität 
* -------------------------------------------------------------------
 */

package de.muenchen.allg.afid;

import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeSet;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XTypeProvider;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.util.XCloseable;

/**
 * The class UnoService is a wrapper for UnoService-Objects provided by the
 * OpenOffice.org API. The main aim of this wrapper is to make java-programs
 * more readable by avoiding the neccessity of UnoRuntime.queryInterface-calls.
 * The wrapper performs the queryInterface-calls and holds the corresponding
 * interface-instances. Temporary variables for singular interface-types are no
 * longer needed in your own java-programs.
 * 
 * The wrapper provides convenience-methods to certain aspects of uno-services
 * like getter and setter for property-values.
 * 
 * The wrapper also increases tranzperency when working with uno-objects. It
 * provides methods to inspect uno-services at runtime and print information
 * about supported-services, implemented interfaces, properties and methods. It
 * makes Uno-Programming in java similar to OOo-Basic programming using tools
 * like the famous XRay-tool.
 */
public class UnoService {
	/**
	 * This field contains the original object of the uno-service.
	 */
	private final Object unoObject;

	/**
	 * This map contains all at least one time required interface-instances of
	 * this uno-service.
	 */
	private Map interfaceMap;

	/**
	 * The constructor creates a new wrapperclass for a given uno-service
	 * object.
	 * 
	 * @param unoObject
	 *            an arbitrary object returned by the OOo-API
	 */
	public UnoService(Object unoObject) {
		this.unoObject = unoObject;
		interfaceMap = new HashMap();
	}

	/**
	 * This method returns the interface-instance of this UnoService for a
	 * specified interface-type. It performs a UnoRuntime.queryInterface call
	 * and caches the result interface-type. If the service doesn't implement
	 * the required interface, the method returns <code>null</code>. Use the
	 * construct <code>if (myUnoService.xAnInterface() != null) ...</code> to
	 * ensure the service really implements the interface.
	 * 
	 * @param ifClass
	 *            interface-type to query for.
	 * @return The interface instance for this uno-service. If the service
	 *         doesn't implement the required interface, the method returns
	 *         <code>null</null>.
	 */
	public Object queryInterface(Class ifClass) {
		if (!interfaceMap.containsKey(ifClass)) {
			Object ifInstance = UnoRuntime.queryInterface(ifClass,
					this.unoObject);
			interfaceMap.put(ifClass, ifInstance);
		}
		return interfaceMap.get(ifClass);
	}

	/***************************************************************************
	 * Methods to inspect the uno-service.
	 **************************************************************************/

	/**
	 * This method returns the implementationName of this unoService-object.
	 * 
	 * @return the implementationName of this unoService-object of the String
	 *         "noneUnoService" if the object is not a proper UnoService.
	 */
	public String getImplememtationName() {
		if (this.xServiceInfo() != null) {
			return this.xServiceInfo().getImplementationName();
		} else {
			return "noneUnoService";
		}
	}

	/**
	 * This method provides information about supported services, implemented
	 * interfaces, supported properties and methods of this uno-service. The
	 * lists of services, implemented interfaces and methods are sorted in
	 * alphabetical ascending order.
	 * 
	 * @return a string containing all the inspection information.
	 */
	public String features() {
		String str = "";
		str += "------------------------------------------\n";
		str += "ImplementationName: "
				+ this.xServiceInfo().getImplementationName() + "\n\n";

		str += dbg_Services();
		str += dbg_SupportedInterfaces();
		str += dbg_Properties();
		str += dbg_SupportedInterfacesAndMethods();
		return str;
	}

	/**
	 * This method prints the inspection-information provides by features() to
	 * System.out.
	 */
	public void printFeatures() {
		System.out.println(features());
	}

	/**
	 * This method returns a sorted list of services supported by this
	 * uno-service.
	 * 
	 * @return a string containig the inspection-information.
	 */
	private String dbg_Services() {
		String str = "Supported UNO-Services: ";
		if (this.xServiceInfo() != null) {
			str += "\n";
			Iterator servicesIterator = getSortedServiceIterator();
			while (servicesIterator.hasNext()) {
				str += "  " + ((String) servicesIterator.next()) + "\n";
			}
		} else {
			str += "none\n";
		}
		return str + "\n";
	}

	/**
	 * This method returns a sorted list of properties supported by this
	 * uno-service.
	 * 
	 * @return a string containig the inspection-information.
	 */
	private String dbg_Properties() {
		String str = "Supported Properties: ";
		if (this.xPropertySet() != null) {
			str += "\n";
			Property[] props = this.xPropertySet().getPropertySetInfo()
					.getProperties();
			for (int i = 0; i < props.length; i++) {
				str += "  " + props[i].Name + " - "
						+ props[i].Type.getZClass().getName() + "\n";
			}
		} else {
			str = "none";
		}
		return str + "\n";
	}

	/**
	 * This method returns a sorted list of implemented interfaces.
	 * 
	 * @return a string containig the inspection-information.
	 */
	public String dbg_SupportedInterfaces() {
		String str = "Supported Interfaces: ";
		if (this.xTypeProvider() != null) {
			str += "\n";
			Iterator typesIterator = getSortedTypesIterator();
			while (typesIterator.hasNext()) {
				str += "  " + ((Type) typesIterator.next()).getTypeName()
						+ "\n";
			}
		} else {
			str += "none";
		}
		return str + "\n";
	}

	/**
	 * This method returns a sorted list of implemented interfaces and their
	 * corresponding methods.
	 * 
	 * @return a string containig the inspection-information.
	 */
	private String dbg_SupportedInterfacesAndMethods() {
		String str = "Supported Interfaces and Methods: ";
		if (this.xTypeProvider() != null) {
			str += "\n";
			Iterator typesIterator = getSortedTypesIterator();
			while (typesIterator.hasNext()) {
				Type type = (Type) typesIterator.next();
				str += "  " + type.getTypeName() + "\n";
				Iterator methodsIterator = getSortedMethodsIterator(type);
				while (methodsIterator.hasNext()) {
					Method m = (Method) methodsIterator.next();
					str += "    - " + m.getName() + "(";
					Class[] par = m.getParameterTypes();
					for (int k = 0; k < par.length; k++) {
						if (k != 0) {
							str += ", ";
						}
						str += par[k].getName();
					}
					str += ") - " + m.getReturnType().getName() + "\n";

				}
			}
		} else {
			str += "none";
		}
		return str + "\n";
	}

	/**
	 * This method returns an Iterator to a sorted list of supported services.
	 */
	private Iterator getSortedServiceIterator() {
		return getSortedArrayIterator(this.xServiceInfo()
				.getSupportedServiceNames(), new Comparator() {
			public int compare(Object arg0, Object arg1) {
				return ((String) arg0).compareTo((String) arg1);
			}

			public boolean equals(Object obj) {
				return this == obj;
			}
		});
	}

	/**
	 * This method returns an Iterator to a sorted list of methods supported by
	 * the specified interface-type.
	 */
	private Iterator getSortedMethodsIterator(Type type) {
		Iterator methodsIterator = getSortedArrayIterator(type.getZClass()
				.getDeclaredMethods(), new Comparator() {
			public int compare(Object arg0, Object arg1) {
				return ((Method) arg0).getName().compareTo(
						((Method) arg1).getName());
			}

			public boolean equals(Object obj) {
				return this == obj;
			}
		});
		return methodsIterator;
	}

	/**
	 * This method returns an Iterator to a sorted List of implemented
	 * interface-types.
	 */
	private Iterator getSortedTypesIterator() {
		Iterator i = getSortedArrayIterator(this.xTypeProvider().getTypes(),
				new Comparator() {
					public int compare(Object arg0, Object arg1) {
						return ((Type) arg0).getTypeName().compareTo(
								((Type) arg1).getTypeName());
					}

					public boolean equals(Object obj) {
						return this == obj;
					}
				});
		return i;
	}

	/**
	 * This generic method sorts an array comparing its elements using a
	 * comperator. The method is needes by the methods
	 * getSortedServiceIterator(), getSortedMethodsIterator(Type type) and
	 * getSortedTypesIterator(),
	 */
	private Iterator getSortedArrayIterator(Object[] array, Comparator c) {
		TreeSet set = new TreeSet(c);
		for (int i = 0; i < array.length; i++) {
			set.add(array[i]);
		}
		return set.iterator();
	}

	/***************************************************************************
	 * convenience-methods
	 **************************************************************************/

	/**
	 * This method returns the property of the uno-service and wrappes the
	 * result in a new UnoService-object.
	 * 
	 * @param p
	 *            the name of the required property
	 * @return the value of the property, wrapped in an UnoService object.
	 * @throws Exception
	 */
	public UnoService getPropertyValue(String p) throws Exception {
		if (this.xPropertySet() != null) {
			return new UnoService(this.xPropertySet().getPropertyValue(p));
		} else {
			throw new Exception(
					"Service doesn't support interface XPropertySet");
		}
	}

	/**
	 * This method sets a property-value of this
	 * 
	 * @param p
	 *            the name of the property to set
	 * @param o
	 *            the value of the property
	 * @throws Exception
	 */
	public void setPropertyValue(String p, Object o) throws Exception {
		if (this.xPropertySet() != null) {
			this.xPropertySet().setPropertyValue(p, o);
		} else {
			throw new Exception(
					"Service doesn't support interface XPropertySet");
		}
	}

	/**
	 * This method states if this servic-object provides a specified
	 * uno-service.
	 * 
	 * @param s
	 *            the fullqualified name of the service.
	 * @return true if the object supports the uno-service, otherwise false.
	 */
	public boolean supportsService(String s) {
		if (this.xServiceInfo() != null) {
			return this.xServiceInfo().supportsService(s);
		} else {
			return false;
		}
	}

	/**
	 * This method returns the original object of this uno-service.
	 * 
	 * @return original object of this uno-service
	 */
	public Object getObject() {
		return unoObject;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see java.lang.Object#toString()
	 */
	public String toString() {
		return unoObject.toString();
	}

	/***************************************************************************
	 * Interface-Access
	 * 
	 * Extend this section if you need to access an interface not in this list.
	 **************************************************************************/

	public XTextDocument xTextDocument() {
		return (XTextDocument) queryInterface(XTextDocument.class);
	}

	public XComponentLoader xComponentLoader() {
		return (XComponentLoader) queryInterface(XComponentLoader.class);
	}

	public XDesktop xDesktop() {
		return (XDesktop) queryInterface(XDesktop.class);
	}

	public XText xText() {
		return (XText) queryInterface(XText.class);
	}

	public XTextRange xTextRange() {
		return (XTextRange) queryInterface(XTextRange.class);
	}

	public XPropertySet xPropertySet() {
		return (XPropertySet) queryInterface(XPropertySet.class);
	}

	public XCloseable xCloseable() {
		return (XCloseable) queryInterface(XCloseable.class);
	}

	public XComponentContext xComponentContext() {
		return (XComponentContext) queryInterface(XComponentContext.class);
	}

	public XMultiComponentFactory xMultiComponentFactory() {
		return (XMultiComponentFactory) queryInterface(XMultiComponentFactory.class);
	}

	public XMultiServiceFactory xMultiServiceFactory() {
		return (XMultiServiceFactory) queryInterface(XMultiServiceFactory.class);
	}

	public XUnoUrlResolver xUnoUrlResolver() {
		return (XUnoUrlResolver) queryInterface(XUnoUrlResolver.class);
	}

	public XServiceInfo xServiceInfo() {
		return (XServiceInfo) queryInterface(XServiceInfo.class);
	}

	public XPropertySetInfo xPropertySetInfo() {
		return (XPropertySetInfo) queryInterface(XPropertySetInfo.class);
	}

	public XInterface xInterface() {
		return (XInterface) queryInterface(XInterface.class);
	}

	public XTypeProvider xTypeProvider() {
		return (XTypeProvider) queryInterface(XTypeProvider.class);
	}

	public XEnumerationAccess xEnumerationAccess() {
		return (XEnumerationAccess) queryInterface(XEnumerationAccess.class);
	}

	public XEnumeration xEnumeration() {
		return (XEnumeration) queryInterface(XEnumeration.class);
	}

	public XTextField xTextField() {
		return (XTextField) queryInterface(XTextField.class);
	}

	public XTextContent xTextContent() {
		return (XTextContent) queryInterface(XTextContent.class);
	}

	public XShape xShape() {
		return (XShape) queryInterface(XShape.class);
	}

	//... add wrapper-methods for your own interfaces here...

}
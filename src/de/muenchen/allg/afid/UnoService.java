/*
 * Ein einfacher Wrapper für uno-services.
 * Copyright (C) 2005 Christoph Lutz
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is furnished
 * to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE. 
 * 
 * Änderungshistorie der Landeshauptstadt München:
 * Alle Änderungen (c) Landeshauptstadt München, alle Rechte vorbehalten
 * 
 * Datum      | Wer | Anderungsgrund
 * -------------------------------------------------------------------
 * 26.04.2005 | BNK | getSimpleName() durch getName() 
 *            |     | ersetzt wg. Java 1.4 Kompatibilität
 * 05.09.2005 | BNK | getImplememtationName() -> getImplementationName() 
 * 09.09.2005 | LUT | +xFilePicker()
 * 09.09.2005 | LUT | BUGFIX: getSortedMethodsIterator holt die Methoden
 *                    jetzt mit getMethods statt mit getDeclaredMethods  
 * 13.09.2005 | LUT | +xFrame()
 * 19.09.2005 | LUT | +xToolkit()
 *                    +xExtendedToolkit()
 * 20.09.2005 | LUT | BUGFIX: introspection wirft NullPointerExceptions in
 *                    getSortedTypesIterator, getSortedServiceIterator,
 *                    getImplementationName
 * 20.09.2005 | LUT | + xWindow()        
 *                    + xWindowPeer()  
 *                    + xComboBox()     
 *                    + xModel()
 *                    + xLayoutManager()
 *                    + xModuleUIConfigurationManagerSupplier()
 *                    + xModuleUIConfigurationManager()
 *                    + create()
 *                    + createWithContext()
 * 30.09.2005 | LUT | + xIndexAccess()
 *                    + xIndexContainer()
 *                    +	xElementAccess()
 *                    +	xDockableWindow()
 *                    +	xListBox()
 *                    +	xVclWindowPeer()
 *                    +	xTextComponent()
 *                    + xModuleManager()
 *                    +	xUIElementSettings()
 *                    +	xConfigManager()
 *                    +	xStringSubstitution()
 *                    +	xSimpleRegistry()
 *                    +	xRegistryKey()
 * 14.10.2005 | LUT | + xBookmarkSupplier()                  
 * 17.10.2005 | LUT | + xNamed()  
 *                    + xDocumentInsertable()                
 *                    + xTextCursor()
 * 18.10.2005 | LUT | + xURLTransformer()
 *                    + xTextRangeCompare()
 * 20.10.2005 | LUT | + xTextFieldSupplier()
 *                    + xEventBroadcaster()
 * 02.11.2005 | LUT | + xComponent()
 * 04.11.2005 | LUT | + xDevice()           
 * 08.11.2005 | LUT | + xDocumentInfoSupplier()           
 *                    + xNameAccess()
 *                    + createWithArguments()
 * 14.11.2005 | LUT | + xModifiable()           
 * -------------------------------------------------------------------
 */

package de.muenchen.allg.afid;

import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeSet;

import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XDevice;
import com.sun.star.awt.XDockableWindow;
import com.sun.star.awt.XExtendedToolkit;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XVclWindowPeer;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.container.XElementAccess;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.XDocumentInfoSupplier;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XConfigManager;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XModuleManager;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XTypeProvider;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.registry.XSimpleRegistry;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextRangeCompare;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIElementSettings;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XStringSubstitution;
import com.sun.star.util.XURLTransformer;

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
 * The wrapper also increases transparency when working with uno-objects. It
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
            Object ifInstance = UnoRuntime.queryInterface(
                ifClass, this.unoObject);
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
    public String getImplementationName() {
        if (this.xServiceInfo() != null) {
            return this.xServiceInfo().getImplementationName();
        } else {
            return "none";
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
        if (unoObject != null) {
            String str = "";
            str += "------------------------------------------\n";
            str += "ImplementationName: " + getImplementationName() + "\n\n";

            str += dbg_Services();
            str += dbg_SupportedInterfaces();
            str += dbg_Properties();
            str += dbg_SupportedInterfacesAndMethods();
            return str;
        } else {
            return "null";
        }
    }

    /**
     * This method prints the inspection-information provides by features() to
     * System.out.
     */
    public void printFeatures() {
        System.out.println(features());
    }

    /**
     * This method shows the features of a service in a MsgBox.
     */
    public void msgboxFeatures() {
        MsgBox.simple("Xray: " + getImplementationName(), features());
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
            Property[] props = this
                .xPropertySet().getPropertySetInfo().getProperties();
            for (int i = 0; i < props.length; i++) {
                str += "  " + props[i].Name + " - "
                        + props[i].Type.getZClass().getName() + "\n";
            }
        } else {
            str += "none\n";
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
        if (this.xServiceInfo() != null) {
            return getSortedArrayIterator(this
                .xServiceInfo().getSupportedServiceNames(), new Comparator() {
                public int compare(Object arg0, Object arg1) {
                    return ((String) arg0).compareTo((String) arg1);
                }

                public boolean equals(Object obj) {
                    return this == obj;
                }
            });
        } else {
            return new Iterator() {

                public void remove() {
                }

                public Object next() {
                    return null;
                }

                public boolean hasNext() {
                    return false;
                }
            };
        }
    }

    /**
     * This method returns an Iterator to a sorted list of methods supported by
     * the specified interface-type.
     */
    private Iterator getSortedMethodsIterator(Type type) {
        Iterator methodsIterator = getSortedArrayIterator(type
            .getZClass().getMethods(), new Comparator() {
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
        if (this.xTypeProvider() != null) {
            return getSortedArrayIterator(
                this.xTypeProvider().getTypes(), new Comparator() {
                    public int compare(Object arg0, Object arg1) {
                        return ((Type) arg0).getTypeName().compareTo(
                            ((Type) arg1).getTypeName());
                    }

                    public boolean equals(Object obj) {
                        return this == obj;
                    }
                });
        } else {
            return new Iterator() {

                public void remove() {
                }

                public Object next() {
                    return null;
                }

                public boolean hasNext() {
                    return false;
                }
            };
        }
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

    /**
     * This method creates and returns a new UnoService using the
     * xMultiServiceFactory of the current unoObject.
     * 
     * @param name
     *            the name of the required service
     * @return The new created UnoService. If the service does not exist or the
     *         unoObject doesn't support xMultiServiceFactory, the result is a
     *         new UnoObject(null).
     * @throws Exception
     */
    public UnoService create(String name) throws Exception {
        if (this.xMultiServiceFactory() != null) {
            return new UnoService(this.xMultiServiceFactory().createInstance(
                name));
        } else {
            return new UnoService(null);
        }
    }

    /**
     * This method creates a new UnoService with arguments using the
     * xMultiServiceFactory of the current unoObject.
     * 
     * @param name
     *            the name of the required service
     * @param args
     *            the used arguments as a array of PropertyValues
     * @return The new created UnoService. If the service does not exist or the
     *         unoObject doesn't support xMultiServiceFactory, the result is a
     *         new UnoObject(null).
     * @throws Exception
     */
    public UnoService create(String name, Object[] args) throws Exception {
        if (this.xMultiServiceFactory() != null) {
            return new UnoService(this
                .xMultiServiceFactory().createInstanceWithArguments(name, args));
        } else {
            return new UnoService(null);
        }
    }

    /**
     * This method creates a new UnoService with a specified context.
     * 
     * @param name
     *            the name of the required service.
     * @param ctx
     *            the context of the new service.
     * @return the new UnoService.
     * @throws Exception
     */
    public static UnoService createWithContext(String name,
            XComponentContext ctx) throws Exception {
        return new UnoService(ctx
            .getServiceManager().createInstanceWithContext(name, ctx));
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

    public XFilePicker xFilePicker() {
        return (XFilePicker) queryInterface(XFilePicker.class);
    }

    public XFrame xFrame() {
        return (XFrame) queryInterface(XFrame.class);
    }

    public XToolkit xToolkit() {
        return (XToolkit) queryInterface(XToolkit.class);
    }

    public XExtendedToolkit xExtendedToolkit() {
        return (XExtendedToolkit) queryInterface(XExtendedToolkit.class);
    }

    public XWindow xWindow() {
        return (XWindow) queryInterface(XWindow.class);
    }

    public XWindowPeer xWindowPeer() {
        return (XWindowPeer) queryInterface(XWindowPeer.class);
    }

    public XComboBox xComboBox() {
        return (XComboBox) queryInterface(XComboBox.class);
    }

    public XModel xModel() {
        return (XModel) queryInterface(XModel.class);
    }

    public XLayoutManager xLayoutManager() {
        return (XLayoutManager) queryInterface(XLayoutManager.class);
    }

    public XModuleUIConfigurationManagerSupplier xModuleUIConfigurationManagerSupplier() {
        return (XModuleUIConfigurationManagerSupplier) queryInterface(XModuleUIConfigurationManagerSupplier.class);
    }

    public XUIConfigurationManager xUIConfigurationManager() {
        return (XUIConfigurationManager) queryInterface(XUIConfigurationManager.class);
    }

    public XIndexAccess xIndexAccess() {
        return (XIndexAccess) queryInterface(XIndexAccess.class);
    }

    public XIndexContainer xIndexContainer() {
        return (XIndexContainer) queryInterface(XIndexContainer.class);
    }

    public XElementAccess xElementAccess() {
        return (XElementAccess) queryInterface(XElementAccess.class);
    }

    public XDockableWindow xDockableWindow() {
        return (XDockableWindow) queryInterface(XDockableWindow.class);
    }

    public XListBox xListBox() {
        return (XListBox) queryInterface(XListBox.class);
    }

    public XVclWindowPeer xVclWindowPeer() {
        return (XVclWindowPeer) queryInterface(XVclWindowPeer.class);
    }

    public XTextComponent xTextComponent() {
        return (XTextComponent) queryInterface(XTextComponent.class);
    }

    public XModuleManager xModuleManager() {
        return (XModuleManager) queryInterface(XModuleManager.class);
    }

    public XUIElementSettings xUIElementSettings() {
        return (XUIElementSettings) queryInterface(XUIElementSettings.class);
    }

    public XConfigManager xConfigManager() {
        return (XConfigManager) queryInterface(XConfigManager.class);
    }

    public XStringSubstitution xStringSubstitution() {
        return (XStringSubstitution) queryInterface(XStringSubstitution.class);
    }

    public XSimpleRegistry xSimpleRegistry() {
        return (XSimpleRegistry) queryInterface(XSimpleRegistry.class);
    }

    public XRegistryKey xRegistryKey() {
        return (XRegistryKey) queryInterface(XRegistryKey.class);
    }

    public XBookmarksSupplier xBookmarksSupplier() {
        return (XBookmarksSupplier) queryInterface(XBookmarksSupplier.class);
    }

    public XNamed xNamed() {
        return (XNamed) queryInterface(XNamed.class);
    }

    public XDocumentInsertable xDocumentInsertable() {
        return (XDocumentInsertable) queryInterface(XDocumentInsertable.class);
    }

    public XTextCursor xTextCursor() {
        return (XTextCursor) queryInterface(XTextCursor.class);
    }

    public XURLTransformer xURLTransformer() {
        return (XURLTransformer) queryInterface(XURLTransformer.class);
    }

    public XTextRangeCompare xTextRangeCompare() {
        return (XTextRangeCompare) queryInterface(XTextRangeCompare.class);
    }

    public XTextFieldsSupplier xTextFieldsSupplier() {
        return (XTextFieldsSupplier) queryInterface(XTextFieldsSupplier.class);
    }

    public XEventBroadcaster xEventBroadcaster() {
        return (XEventBroadcaster) queryInterface(XEventBroadcaster.class);
    }

    public XComponent xComponent() {
        return (XComponent) queryInterface(XComponent.class);
    }

    public XDevice xDevice() {
        return (XDevice) queryInterface(XDevice.class);
    }

    public XDispatch xDispatch() {
        return (XDispatch) queryInterface(XDispatch.class);
    }

    public XDocumentInfoSupplier xDocumentInfoSupplier() {
        return (XDocumentInfoSupplier) queryInterface(XDocumentInfoSupplier.class);
    }

    public XNameAccess xNameAccess() {
        return (XNameAccess) queryInterface(XNameAccess.class);
    }

    public XModifiable xModifiable() {
        return (XModifiable) queryInterface(XModifiable.class);
    }

    // ... add wrapper-methods for your own interfaces here...

}
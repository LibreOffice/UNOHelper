/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2022 Landeshauptstadt München
 * %%
 * Licensed under the EUPL, Version 1.1 or – as soon they will be
 * approved by the European Commission - subsequent versions of the
 * EUPL (the "Licence");
 *
 * You may not use this work except in compliance with the Licence.
 * You may obtain a copy of the Licence at:
 *
 * http://ec.europa.eu/idabc/eupl5
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the Licence is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the Licence for the specific language governing permissions and
 * limitations under the Licence.
 * #L%
 */
package de.muenchen.allg.document.text;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.style.XStyle;
import com.sun.star.text.XTextDocument;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoDictionary;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.util.UnoService;

/**
 * Service for creating and accessing styles.
 */
public class StyleService
{

  public static final String PARAGRAPH_STYLES = "ParagraphStyles";

  public static final String CHARACTER_STYLES = "CharacterStyles";

  public static final String PAGE_STYLES = "PageStyles";

  private StyleService()
  {
  }

  /**
   * Get the paragraph style with a given name.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the paragraph style.
   * @return The style with the name.
   * @throws UnoHelperException
   *           Can't get paragraph style.
   */
  public static XStyle getParagraphStyle(XTextDocument doc, String name) throws UnoHelperException
  {
    try
    {
      XNameContainer pss = getStyleContainer(doc, PARAGRAPH_STYLES);
      return UNO.XStyle(pss.getByName(name));
    } catch (NullPointerException | NoSuchElementException | WrappedTargetException e)
    {
      throw new UnoHelperException("Can't get paragraph style", e);
    }
  }

  /**
   * Does the same as {@link #getParagraphStyle(XTextDocument, String)} but returns a default value
   * instead of throwing a exception.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the sty.e
   * @param def
   *          The default value.
   * @return The style or the default value.
   */
  public static XStyle getParagraphStyle(XTextDocument doc, String name, XStyle def)
  {
    try
    {
      return getParagraphStyle(doc, name);
    } catch (UnoHelperException e)
    {
      return def;
    }
  }

  /**
   * Creates a new paragraph style.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the paragraph style.
   * @param parentStyleName
   *          The name of the parent paragraph style or null if there is no parent, which defaults
   *          to "Standard"
   * @return The style or null if it cloudn't be created.
   * @throws UnoHelperException
   *           Can't create the style.
   */
  public static XStyle createParagraphStyle(XTextDocument doc, String name, String parentStyleName)
      throws UnoHelperException
  {
    return createStyle(doc, getStyleContainer(doc, PARAGRAPH_STYLES), UnoService.CSS_STYLE_PARAGRAPH_STYLE, name,
        parentStyleName);
  }

  /**
   * Get the character style with a given name.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the character style.
   * @return The style with the name.
   * @throws UnoHelperException
   *           Can't get character style.
   */
  public static XStyle getCharacterStyle(XTextDocument doc, String name) throws UnoHelperException
  {
    try
    {
      XNameContainer pss = getStyleContainer(doc, CHARACTER_STYLES);
      return UNO.XStyle(pss.getByName(name));
    } catch (NullPointerException | NoSuchElementException | WrappedTargetException e)
    {
      throw new UnoHelperException("Can't get character style.", e);
    }
  }

  /**
   * Does the same as {@link #getCharacterStyle(XTextDocument, String)} but returns a default value
   * instead of throwing a exception.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the sty.e
   * @param def
   *          The default value.
   * @return The style or the default value.
   */
  public static XStyle getCharacterStyle(XTextDocument doc, String name, XStyle def)
  {
    try
    {
      return getCharacterStyle(doc, name);
    } catch (UnoHelperException e)
    {
      return def;
    }
  }

  /**
   * Creates a new character style.
   *
   * @param doc
   *          The document.
   * @param name
   *          The name of the character style.
   * @param parentStyleName
   *          The name of the parent paragraph style or null if there is no parent, which defaults
   *          to "Standard"
   * @return The style or null if it cloudn't be created.
   * @throws UnoHelperException
   *           Can't create the style.
   */
  public static XStyle createCharacterStyle(XTextDocument doc, String name, String parentStyleName)
      throws UnoHelperException
  {

    return createStyle(doc, getStyleContainer(doc, CHARACTER_STYLES), UnoService.CSS_STYLE_CHARACTER_STYLE, name,
        parentStyleName);
  }

  /**
   * Create a new style.
   * 
   * @param doc
   *          The document.
   * @param styles
   *          The style container of the document.
   * @param styleType
   *          The type of style.
   * @param name
   *          The name of the style.
   * @param parentStyleName
   *          The parent style.
   * @return The new style
   * @throws UnoHelperException
   *           Can't create the style.
   */
  private static XStyle createStyle(XTextDocument doc, XNameContainer styles, String styleType, String name,
      String parentStyleName) throws UnoHelperException
  {
    try
    {

      XStyle style = UNO.XStyle(UnoService.createService(styleType, doc));
      styles.insertByName(name, style);
      if (style != null && parentStyleName != null)
      {
        style.setParentStyle(parentStyleName);
      }
      return UNO.XStyle(styles.getByName(name));
    } catch (Exception e)
    {
      throw new UnoHelperException("Can't create style", e);
    }
  }

  /**
   * Get the styles of the document.
   *
   * @param doc
   *          The document.
   * @param containerName
   *          The type of styles ({@link #CHARACTER_STYLES}, {@link #PARAGRAPH_STYLES})
   * @return The container of the styles or null.
   * @throws UnoHelperException
   *           A style container with this name doesn't exist.
   */
  public static XNameContainer getStyleContainer(XTextDocument doc, String containerName) throws UnoHelperException
  {
    try
    {
      return UnoDictionary.create(UNO.XStyleFamiliesSupplier(doc).getStyleFamilies(), XNameContainer.class)
          .get(containerName);
    } catch (NullPointerException e)
    {
      throw new UnoHelperException("No such style container", e);
    }
  }

}

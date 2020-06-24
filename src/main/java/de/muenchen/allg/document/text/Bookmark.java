/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2020 Landeshauptstadt München
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
import com.sun.star.container.XNamed;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.AnyConverter;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoCollection;
import de.muenchen.allg.afid.UnoDictionary;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.util.UnoProperty;
import de.muenchen.allg.util.UnoService;
/**
 * Helper for working with LibreOffice book marks.
 */
public class Bookmark
{

  /**
   * Error message if the book mark can't be accessed.
   */
  private static final String ACCESS_ERROR_MESSAGE = "Can't access the book mark.";

  /**
   * The name of book marks, which were removed.
   */
  public static final String BROKEN = "WM(CMD'bookmarkBroken')";

  /**
   * The name of the book mark.
   */
  private String name;

  /**
   * The document containing the book mark..
   */
  private XTextDocument document;

  /**
   * Access an existing book mark in the document.
   *
   * @param name
   *          The name of the book mark.
   * @param doc
   *          The document.
   * @throws UnoHelperException
   *           A book mark with this name doesn't exist.
   */
  public Bookmark(String name, XBookmarksSupplier doc) throws UnoHelperException
  {
    this.document = UNO.XTextDocument(doc);
    this.name = name;
    XTextContent bookmark = getBookmarkService(name);
    if (bookmark == null)
    {
      throw new UnoHelperException(String.format("Bookmark '%s' existiert nicht.", name));
    }
  }

  /**
   * Access an existing book mark in the document.
   *
   * @param bookmark
   *          The name of the book mark.
   * @param doc
   *          The document.
   */
  public Bookmark(XNamed bookmark, XTextDocument doc)
  {
    this.document = doc;
    this.name = bookmark.getName();
  }

  /**
   * Create a new book mark at the given position. If the position has no dimension, a collapsed
   * book mark is created.
   *
   * @param name
   *          The name of the book mark.
   * @param doc
   *          The document.
   * @param range
   *          The position of the book mark.
   * @throws UnoHelperException
   *           Can't create the book mark.
   */
  public Bookmark(String name, XTextDocument doc, XTextRange range) throws UnoHelperException
  {
    this.document = doc;
    this.name = name;

    XTextContent bookmark = null;
    try
    {
      bookmark = UNO.XTextContent(UnoService.createService(UnoService.CSS_TEXT_BOOKMARK, document));
    } catch (Exception e)
    {
      throw new UnoHelperException("Can't create the book mark.", e);
    }

    if (UNO.XNamed(bookmark) != null)
    {
      UNO.XNamed(bookmark).setName(name);
    }

    // add book mark to the document
    if (document != null && bookmark != null && range != null)
    {
      try
      {
        // use text cursor instead of text range because they have more functionalty
        XTextCursor cursor = range.getText().createTextCursorByRange(range);
        range.getText().insertTextContent(cursor, bookmark, !cursor.isCollapsed());
        this.name = UNO.XNamed(bookmark).getName();
      } catch (IllegalArgumentException e)
      {
        throw new UnoHelperException("Can't create the book mark.", e);
      }
    }
  }

  public String getName()
  {
    return name;
  }

  public XTextDocument getDocument()
  {
    return document;
  }

  /**
   * Select the text under the book mark.
   * 
   * @throws UnoHelperException
   *           Can't access the book mark.
   */
  public void select() throws UnoHelperException
  {
    XTextContent bm = getBookmarkService(getName());
    if (bm != null)
    {
      XTextRange anchor = bm.getAnchor();
      XTextRange cursor = anchor.getText().createTextCursorByRange(anchor);
      UNO.XTextViewCursorSupplier(UNO.XModel(document).getCurrentController()).getViewCursor().gotoRange(cursor, false);
    }
  }

  /**
   * Rename this book mark. If the given name is already remitted, a number is added at the end.
   *
   * @param newName
   *          The requested name of the book mark.
   * @return The new name or {@link #BROKEN} if this book mark doesn't exist anymore.
   */
  public String rename(String newName)
  {
    UnoDictionary<XTextContent> bookmarks = UnoDictionary
        .create(UNO.XBookmarksSupplier(document).getBookmarks(),
        XTextContent.class);

    // this book mark has already the request name
    if (name.equals(newName))
    {
      if (!bookmarks.containsKey(name))
      {
        name = BROKEN;
      }
      return name;
    }

    // add number if names is already remitted
    if (bookmarks.containsKey(newName))
    {
      int count = 1;
      while (bookmarks.containsKey(newName + count))
      {
        ++count;
      }
      newName = newName + count;
    }

    XNamed bm = UNO.XNamed(bookmarks.get(name));
    if (bm != null)
    {
      bm.setName(newName);
      name = bm.getName();
    } else
    {
      name = BROKEN;
    }

    return name;
  }

  /**
   * Move the book mark to a new text range.
   *
   * @param xTextRange
   *          The new range.
   * @throws UnoHelperException
   *           Can't move the book mark.
   */
  public void rerangeBookmark(XTextRange xTextRange) throws UnoHelperException
  {
    // create new book mark with old name and new dimension
    try
    {
      XTextContent bookmark = UNO.XTextContent(UnoService.createService(UnoService.CSS_TEXT_BOOKMARK, document));
      // delete old book mark
      remove();
      UNO.XNamed(bookmark).setName(name);
      xTextRange.getText().insertTextContent(xTextRange, bookmark, true);
    } catch (Exception e)
    {
      throw new UnoHelperException("Can't move the book mark.", e);
    }
  }

  /**
   * Get a {@link XTextCursor} of the book mark. This text cursor is more robust than the anchor of
   * {@link #getAnchor()}.
   *
   * @return A text cursor over the range where this book mark is anchoring or null if the book mark
   *         doesn't exist anymore.
   * @throws UnoHelperException
   *           Can't access the book mark.
   */
  public XTextCursor getTextCursor() throws UnoHelperException
  {
    /**
     * work around for https://bz.apache.org/ooo/show_bug.cgi?id=67869: wrong anchor for book marks
     * in tables. A text cursor is more robust.
     */
    XTextRange range = getAnchor();
    if (range != null && range.getText() != null)
    {
      return range.getText().createTextCursorByRange(range);
    }
    return null;
  }

  /**
   * Get the {@link XTextRange} of this book mark. You may use {@link #getTextCursor()} because of
   * OOo-Issue #67869.
   *
   * @return The text range where this book mark is anchoring or null if the book mark doesn't exist
   *         anymore.
   * @throws UnoHelperException
   *           Can't access the book mark.
   */
  public XTextRange getAnchor() throws UnoHelperException
  {
    XBookmarksSupplier supp = UNO.XBookmarksSupplier(document);
    try
    {
      return UNO.XTextContent(supp.getBookmarks().getByName(name)).getAnchor();
    } catch (NoSuchElementException | WrappedTargetException x)
    {
      throw new UnoHelperException(ACCESS_ERROR_MESSAGE, x);
    }
  }

  /**
   * Delete this book mark.
   * 
   * @throws UnoHelperException
   *           Can't delete the book mark.
   */
  public void remove() throws UnoHelperException
  {
    XTextContent bookmark = getBookmarkService(name);
    if (bookmark != null)
    {
      try
      {
        XTextRange range = bookmark.getAnchor();
        range.getText().removeTextContent(bookmark);
      } catch (NoSuchElementException e)
      {
        throw new UnoHelperException("Can't delete the book mark.", e);
      }
    }
  }

  /**
   * Get the value of the property "IsCollapsed". Therefore the {@link XTextCursor} is iterated.
   *
   * @return True, if the property "IsCollapsed" exists and is true., false otherwise.
   * @throws UnoHelperException
   *           Can't access the book mark.
   */
  public boolean isCollapsed() throws UnoHelperException
  {
    XTextRange anchor = getTextCursor();
    if (anchor == null)
    {
      return false;
    }
    try
    {
      Object par = UNO.XEnumerationAccess(anchor).createEnumeration().nextElement();
      UnoCollection<Object> collection = UnoCollection.getCollection(par, Object.class);
      for (Object content : collection)
      {
        String tpt = "" + UnoProperty.getProperty(content, UnoProperty.TEXT_PROTION_TYPE);
        if (UnoProperty.BOOKMARK.equals(tpt))
        {
          XNamed bm = UNO.XNamed(UnoProperty.getProperty(content, UnoProperty.BOOKMARK));
          if (bm != null && name.equals(bm.getName()))
          {
            return AnyConverter.toBoolean(UnoProperty.getProperty(content, UnoProperty.IS_COLLAPSED));
          }
        }
      }
    } catch (IllegalArgumentException | NoSuchElementException | WrappedTargetException e)
    {
      throw new UnoHelperException(ACCESS_ERROR_MESSAGE, e);
    }
    return false;
  }

  /**
   * Expand a book mark. It has no dimension.
   * 
   * @throws UnoHelperException
   *           Can't decollapse the book mark.
   */
  public void decollapseBookmark() throws UnoHelperException
  {
    XTextRange range = getAnchor();
    if (range == null || !isCollapsed())
    {
      return;
    }

    XTextCursor cursor = range.getText().createTextCursorByRange(range);

    // create new book mark with old name and new range
    try
    {
      XTextContent bookmark = UNO.XTextContent(UnoService.createService(UnoService.CSS_TEXT_BOOKMARK, document));
      remove();
      UNO.XNamed(bookmark).setName(name);
      cursor.getText().insertString(cursor, ".", true);
      cursor.getText().insertTextContent(cursor, bookmark, true);
    } catch (Exception e)
    {
      throw new UnoHelperException("Can't decollapse the book mark.", e);
    }
  }

  /**
   * Collapse a book mark. Its content isn't deleted. The new book mark is right in front of the
   * content.
   * 
   * @throws UnoHelperException
   *           Can't collapse the book mark.
   */
  public void collapseBookmark() throws UnoHelperException
  {
    XTextRange range = getAnchor();

    // do nothing, if book mark doesn't exist or is already collapsed
    if (range == null || isCollapsed())
    {
      return;
    }


    // create new collapsed book mark with old name
    try
    {
      XTextContent bookmark = UNO.XTextContent(UnoService.createService(UnoService.CSS_TEXT_BOOKMARK, document));
      remove();
      UNO.XNamed(bookmark).setName(name);
      range.getText().insertTextContent(range.getStart(), bookmark, false);
    } catch (Exception e)
    {
      throw new UnoHelperException("Can't collapse the book mark.", e);
    }
  }

  @Override
  public boolean equals(Object b)
  {
    if (b == null)
    {
      return false;
    }

    if (this.getClass() != b.getClass())
    {
      return false;
    }

    try
    {
      return name.equals(((Bookmark) b).name);
    } catch (java.lang.Exception e)
    {
      return false;
    }
  }

  @Override
  public int hashCode()
  {
    return name.hashCode();
  }

  @Override
  public String toString()
  {
    return "Bookmark[" + getName() + "]";
  }

  /**
   * Get a book mark. Before each access to the book mark it should be collect from the document. It
   * may has been removed in the mean time.
   *
   * @param name
   *          The name of the book mark.
   * @return The text content of the book mark.
   * @throws UnoHelperException
   *           Can't access the book mark.
   */
  private XTextContent getBookmarkService(String name) throws UnoHelperException
  {
    try
    {
      return UNO.XTextContent(UNO.XBookmarksSupplier(document).getBookmarks().getByName(name));
    } catch (NullPointerException | WrappedTargetException | NoSuchElementException e)
    {
      throw new UnoHelperException(ACCESS_ERROR_MESSAGE, e);
    }
  }

}

package de.muenchen.allg.ooo;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNamed;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.afid.UnoHelperRuntimeException;
import de.muenchen.allg.afid.UnoIterator;
import de.muenchen.allg.util.UnoProperty;

/**
 * Helper for working with {@link XTextDocument}s.
 */
public class TextDocument
{
  private TextDocument()
  {
    // nothing to do
  }

  /**
   * Delete the first paragraph at the cursor position.
   *
   * @param range
   *          The cursor position.
   * @throws UnoHelperRuntimeException
   *           Paragraph can't be deleted.
   */
  public static void deleteParagraph(XTextRange range)
  {
    XEnumerationAccess access = UNO.XEnumerationAccess(range);
    if (access != null)
    {
      // get first paragraph
      XTextContent paragraph = new UnoIterator<XTextContent>(access, XTextContent.class).next();

      if (paragraph != null)
      {
        try
        {
          // If the range has only one paragraph it can't be deleted by removeTextContent. So we
          // only overwrite the content with an empty string.
          paragraph.getAnchor().setString("");
          range.getText().removeTextContent(paragraph);
        } catch (NoSuchElementException e)
        {
          throw new UnoHelperRuntimeException(e);
        }
      }
    }
  }

  /**
   * Get the names of all the book marks in the range matching the regular expression. Also names of
   * collapsed book marks are returned. Only book marks which are fully in the range are included.
   *
   * In some special cases like tables this method may return more names than in the text range.
   * 
   * @param regex
   *          A regular expression for the names of the book marks.
   * @param range
   *          The range in which we are looking for book marks.
   * @return All names of book marks in the range matching the regular expression.
   */
  public static Set<String> getBookmarkNamesMatching(Pattern regex, XTextRange range)
  {
    HashSet<String> found = new HashSet<>();
    HashSet<String> started = new HashSet<>();

    UNO.forEachTextPortionInRange(range, o -> {
      try
      {
        XNamed bookmark = UNO.XNamed(UnoProperty.getProperty(o, UnoProperty.BOOKMARK));
        String name = (bookmark != null) ? bookmark.getName() : "";

        if (regex.matcher(name).matches())
        {
          if (Boolean.TRUE.equals(UnoProperty.getProperty(o, UnoProperty.IS_START)))
          {
            if (Boolean.TRUE.equals(UnoProperty.getProperty(o, UnoProperty.IS_COLLAPSED)))
            {
              found.add(name);
            } else
            {
              started.add(name);
            }
          } else if (started.contains(name))
          {
            found.add(name);
          }
        }
      } catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });

    return found;
  }

  /**
   * Get a list of all book marks in a text range.
   *
   * @param range
   *          The text range.
   * @return The list of book marks.
   *
   * @throws UnoHelperRuntimeException
   *           Can't access a book mark.
   */
  public static List<XNamed> getBookmarkByTextRange(XTextRange range)
  {
    ArrayList<XNamed> bookmarks = new ArrayList<>();

    UNO.forEachTextPortionInRange(range, o -> {
      try
      {
        bookmarks.add(UNO.XNamed(UnoProperty.getProperty(o, UnoProperty.BOOKMARK)));
      } catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });

    return bookmarks;
  }
}

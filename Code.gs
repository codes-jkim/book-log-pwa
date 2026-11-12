/**
 * Book Log PWA - Google Apps Script Backend
 * Parses JSON-LD / OG metadata and appends formatted book entries to Google Sheets.
 */

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const SECRET_TOKEN = props.getProperty("SECRET_TOKEN");

    const data = JSON.parse(e.postData.contents);
    const { url, quote, review, token, date } = data;

    // 1. Verify authorization token
    if (token !== SECRET_TOKEN) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: "unauthorized" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Fetch book metadata (JSON-LD first, OG meta tags fallback)
    const bookInfo = fetchBookMetaFromUrl(url);

    // 3. Access current active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];

    const recordDate = date || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");

    // Find next available row based on column B
    const lastRow = sheet.getRange("B:B").getValues().filter(String).length;
    const targetRow = lastRow + 1;

    const imageFormula = bookInfo.image ? `=IMAGE("${bookInfo.image}")` : "";

    // Columns from B: [Date, Cover, Title, Author, Quote, Review]
    const rowData = [
      recordDate,
      imageFormula,
      bookInfo.title,
      bookInfo.author,
      quote,
      review
    ];

    sheet.getRange(targetRow, 2, 1, rowData.length).setValues([rowData]);
    sheet.setRowHeight(targetRow, 140)

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        title: bookInfo.title,
        author: bookInfo.author
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function fetchBookMetaFromUrl(targetUrl) {
  const options = {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
    }
  };

  try {
    const response = UrlFetchApp.fetch(targetUrl, options);
    const html = response.getContentText("UTF-8");

    let title = "";
    let author = "";
    let image = "";

    // 1. Try Schema.org JSON-LD (<script type="application/ld+json">) first
    const ldJsonRegex = /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
    let ldMatch;
    while ((ldMatch = ldJsonRegex.exec(html)) !== null) {
      try {
        const parsed = JSON.parse(ldMatch[1].trim());
        const items = Array.isArray(parsed) ? parsed : [parsed];
        const bookObj = items.find(item => item["@type"] === "Book");

        if (bookObj) {
          if (bookObj.name) title = bookObj.name;
          if (bookObj.image) {
            image = Array.isArray(bookObj.image) ? bookObj.image[0] : bookObj.image;
          }
          if (bookObj.author) {
            if (Array.isArray(bookObj.author)) {
              author = bookObj.author.map(a => a.name || a).join(", ");
            } else if (typeof bookObj.author === "object") {
              author = bookObj.author.name || "";
            } else {
              author = String(bookObj.author);
            }
          }
          break;
        }
      } catch (jsonErr) {
        // Ignore malformed JSON-LD blocks and continue
      }
    }

    // 2. Fallback to Meta tags if any field is still empty
    if (!title) {
      const titleMatch = html.match(/<meta[^>]*property=["']og:title["'][^>]*content=["']([^"']*)["']/i)
                      || html.match(/<meta[^>]*content=["']([^"']*)["'][^>]*property=["']og:title["']/i);
      if (titleMatch) {
        // Clean trailing " | 시리즈 | 저자" or " - 알라딘" if needed
        title = titleMatch[1].split("|")[0].trim();
      }
    }

    if (!author) {
      const authorMatch = html.match(/<meta[^>]*(?:property|name)=["'](?:og:author|author)["'][^>]*content=["']([^"']*)["']/i)
                       || html.match(/<meta[^>]*content=["']([^"']*)["'][^>]*(?:property|name)=["'](?:og:author|author)["']/i);
      if (authorMatch) author = authorMatch[1].trim();
    }

    if (!image) {
      const imageMatch = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']*)["']/i)
                      || html.match(/<meta[^>]*content=["']([^"']*)["'][^>]*property=["']og:image["']/i);
      if (imageMatch) image = imageMatch[1].trim();
    }

    // Decode common HTML entities
    title = decodeEntities(title);
    author = decodeEntities(author);

    return {
      title: title || "Unknown Title",
      author: author || "Unknown Author",
      image: image || ""
    };
  } catch (e) {
    Logger.log("Scraping Error: " + e.message);
    return {
      title: "URL Parsing Failed",
      author: "",
      image: ""
    };
  }
}

function decodeEntities(str) {
  if (!str) return "";
  return str.replace(/&amp;/g, "&")
            .replace(/&lt;/g, "<")
            .replace(/&gt;/g, ">")
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'");
}


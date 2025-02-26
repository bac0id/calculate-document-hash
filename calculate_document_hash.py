import uno
import hashlib
from com.sun.star.script.provider import XScriptContext

def calculate_document_hash():
    """
    Calculate the SHA-256 hash of the current Writer document and insert it into the document.
    """

    # Get the current document model, use XSCRIPTCONTEXT to get the desktop
    desktop = XSCRIPTCONTEXT.getDesktop()
    document = desktop.getCurrentComponent()

    # Check if it is a Writer document
    if document is None or not document.supportsService("com.sun.star.text.TextDocument"):
        print("Please run this macro in a LibreOffice Writer document.")
        return

    # Get the document text content
    text = document.Text.getString()
    text_bytes = text.encode('utf-8') # Encoded as bytes to calculate hash

    # Calculate SHA-256 hash value
    sha256_hash = hashlib.sha256(text_bytes).hexdigest()

    # Get the current cursor position
    view_cursor = document.getCurrentController().getViewCursor()
    text_cursor = document.Text.createTextCursorByRange(view_cursor)
    # Insert the hash value into the document
    text_cursor.String = f"SHA-256 Hash: {sha256_hash}"

    print("Document hash value has been successfully written to the document.")


g_exportedScripts = (calculate_document_hash,) # Export macro

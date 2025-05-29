import os
import sys
from doc_transformation.docx_processor import run_document_processing

def update_status(message):
    """Simple status callback for testing."""
    print(message)

def main():
    """Test the document processing with track changes."""
    if len(sys.argv) != 3:
        print("Usage: python test_track_changes.py <english_docx_path> <chinese_docx_path>")
        return
    
    english_path = sys.argv[1]
    chinese_path = sys.argv[2]
    
    # Create output path in the same directory as the Chinese document
    output_dir = os.path.dirname(chinese_path)
    output_filename = os.path.splitext(os.path.basename(chinese_path))[0] + "_updated_with_track_changes.docx"
    output_path = os.path.join(output_dir, output_filename)
    
    print(f"Processing documents:")
    print(f"English document: {english_path}")
    print(f"Chinese document: {chinese_path}")
    print(f"Output will be saved to: {output_path}")
    
    # Run the document processing
    run_document_processing(english_path, chinese_path, output_path, update_status)
    
    print("\nProcessing complete!")
    print(f"Please open {output_path} in Microsoft Word to verify track changes.")

if __name__ == "__main__":
    main()

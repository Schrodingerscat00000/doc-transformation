# Document Track Changes Processor

This tool processes English documents with tracked changes and applies those changes to a corresponding Chinese document, preserving the track changes markup.

## Key Features

- Identifies tracked changes (insertions and deletions) in English documents
- Locates corresponding text segments in Chinese documents using LLM-based alignment
- Applies the changes to the Chinese document with proper track changes markup
- Preserves the document structure and formatting

## How It Works

The tool uses a hybrid approach combining traditional document processing techniques with LLM capabilities:

1. **Change Extraction**: Parses the English DOCX file to identify all tracked changes
2. **Text Alignment**: Uses the LLM to find corresponding paragraphs between languages
3. **Translation**: Translates inserted text from English to Chinese
4. **Change Application**: Applies the changes to the Chinese document with proper track changes markup

## Usage

### GUI Method

1. Run the application:
   ```
   python -m doc_transformation.app
   ```

2. Select the English document with tracked changes
3. Select the original Chinese document
4. Click "Start Processing"
5. The output will be saved in the same directory as the Chinese document with "_updated_v2" suffix

### Command Line Method

Use the test script for command-line processing:

```
python test_track_changes.py <english_docx_path> <chinese_docx_path>
```

The output will be saved in the same directory as the Chinese document with "_updated_with_track_changes" suffix.

## Requirements

- Python 3.7+
- Dependencies listed in requirements.txt
- Ollama server running locally with the deepseek-r1:1.5b model loaded

## Implementation Details

### Track Changes Preservation

The tool properly preserves track changes by:

1. For insertions:
   - Creating proper `<w:ins>` elements with author, date, and ID attributes
   - Splitting existing runs at the insertion point
   - Preserving text formatting

2. For deletions:
   - Creating proper `<w:del>` elements with author, date, and ID attributes
   - Using `<w:delText>` elements to mark deleted text
   - Handling deletions that span multiple runs

### LLM Integration

The tool uses the deepseek-r1:1.5b model for:

1. Paragraph alignment between languages
2. Translation of inserted text
3. Identification of text to be deleted
4. Finding optimal insertion positions

## Troubleshooting

If the output document doesn't show track changes:

1. Ensure Microsoft Word is set to show tracked changes (Review tab > Track Changes)
2. Check that the "Show markup" option is enabled
3. Verify that both insertions and deletions are selected in the markup options

If the alignment is incorrect:

1. Try using a larger LLM model for better accuracy
2. Consider pre-processing the documents to improve alignment

## License

This project is licensed under the MIT License - see the LICENSE file for details.

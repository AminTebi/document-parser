# Document Parser Script

🚀 **Just built something cool!** 

I created a Python script that automatically extracts hierarchical content from Word documents and converts it into structured JSON format. Perfect for anyone working with structured documents!

## What it does:
✅ Reads Word documents (.docx files)
✅ Extracts Heading 1 → Heading 2 → Content structure
✅ Automatically detects bullet points vs single text
✅ Converts multi-line content into proper JSON arrays
✅ Handles Persian/Arabic text perfectly

## Example output:
```json
{
  "Heading 1": {
    "Heading 2": [
      "Bullet point 1",
      "Bullet point 2", 
      "Bullet point 3"
    ],
    "Another Heading 2": "Single line content"
  }
}
```

## Why this matters:
- **Automation**: No more manual copy-pasting from documents
- **Data Structure**: Clean, structured JSON ready for APIs/databases
- **Multilingual**: Works with RTL languages like Persian
- **Smart Parsing**: Automatically handles different content types

Perfect for content management, documentation systems, or any project that needs to extract structured data from Word documents.

Built with Python + python-docx library. Clean, efficient, and ready to use! 

#Python #Automation #DataExtraction #DocumentProcessing #JSON #TechTools #Productivity

---

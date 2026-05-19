with open("word_document_server/tools/live_tools.py", "r") as f:
    content = f.read()

old = """        # Re-fetch doc from app.Documents to get a fresh COM reference
        # Use Item() to avoid stale references from iteration
        fresh_doc = None
        target_name = filename.lower() if filename else None
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents.Item(i)
            if target_name and d.Name.lower() == target_name:
                fresh_doc = d
                break
        if not fresh_doc:
            return json.dumps({"error": f\"Document '{filename}' not found in open documents\"})
        name = fresh_doc.Name
        # Pass SaveChanges as keyword arg to avoid COM dispatch issues
        fresh_doc.Close(SaveChanges=save_flag)
        # If no documents left, quit Word gracefully
        if app.Documents.Count == 0:
            app.Quit()
        return json.dumps({
            "success": True,
            "closed_document": name,
            "message": f\"Closed '{name}' (save_changes={save_changes})\",
        }, ensure_ascii=False)"""

new = """        # Re-fetch doc from app.Documents to get a fresh COM reference
        # Use Item() to avoid stale references from iteration
        fresh_doc = None
        target_name = filename.lower() if filename else None
        for i in range(1, app.Documents.Count + 1):
            d = app.Documents.Item(i)
            if target_name and d.Name.lower() == target_name:
                fresh_doc = d
                break
        if not fresh_doc:
            return json.dumps({"error": f\"Document '{filename}' not found in open documents\"})
        name = fresh_doc.Name
        # Pass SaveChanges as keyword arg to avoid COM dispatch issues
        fresh_doc.Close(SaveChanges=save_flag)
        # If no documents left, quit Word gracefully
        if app.Documents.Count == 0:
            app.Quit()
        return json.dumps({
            "success": True,
            "closed_document": name,
            "message": f\"Closed '{name}' (save_changes={save_changes})\",
        }, ensure_ascii=False)"""

if old in content:
    content = content.replace(old, new)
    print("Replaced with keyword arg fix")
else:
    print("Could not find old block")

with open("word_document_server/tools/live_tools.py", "w") as f:
    f.write(content)
print("Done")

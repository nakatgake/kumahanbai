import asyncio
import os
from main import generate_quote_pdf, generate_delivery_note_pdf, generate_invoice_pdf
from database import SessionLocal

async def main():
    db = SessionLocal()
    artifact_dir = r"C:\Users\nakatake\.gemini\antigravity\brain\b184d9ba-1b21-4c57-936f-3c33dc28d7e4"
    
    try:
        # Quote
        res = await generate_quote_pdf(1, db)
        if hasattr(res, 'body_iterator'):
            with open(os.path.join(artifact_dir, "direct_quote.pdf"), "wb") as f:
                async for chunk in res.body_iterator:
                    if isinstance(chunk, bytes):
                        f.write(chunk)
                    elif isinstance(chunk, str):
                        f.write(chunk.encode())
        
        # Delivery Note
        res = await generate_delivery_note_pdf(1, db)
        if hasattr(res, 'body_iterator'):
            with open(os.path.join(artifact_dir, "direct_delivery.pdf"), "wb") as f:
                async for chunk in res.body_iterator:
                    if isinstance(chunk, bytes):
                        f.write(chunk)
                    elif isinstance(chunk, str):
                        f.write(chunk.encode())
        
        # Invoice
        res = await generate_invoice_pdf(1, db)
        if hasattr(res, 'body_iterator'):
            with open(os.path.join(artifact_dir, "direct_invoice.pdf"), "wb") as f:
                async for chunk in res.body_iterator:
                    if isinstance(chunk, bytes):
                        f.write(chunk)
                    elif isinstance(chunk, str):
                        f.write(chunk.encode())
                        
        print("Generated direct_quote.pdf, direct_delivery.pdf, direct_invoice.pdf")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        db.close()

if __name__ == "__main__":
    asyncio.run(main())

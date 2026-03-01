import os
import sys

# Attempt to import playwright, if it fails, provide a clear error message.
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright is not installed. Please run: pip install playwright && playwright install")
    sys.exit(1)

def export_dashboards():
    base_dir = r"c:\Users\yassi\OneDrive\Bureau\bilan"
    output_dir = os.path.join(base_dir, "Images_Dashboards")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    html_files = [
        "Tableau_de_Bord_FSSM.html",
        "Tableau_de_Bord_Filiere_Digitale.html",
        "Bilan_Detaille_Infrastructures.html",
        "Tableau_de_Bord_Langues_Power_Skills.html"
    ]
    
    with sync_playwright() as p:
        # Launch headless browser
        browser = p.chromium.launch(headless=True)
        # Use a high-res viewport for good quality images
        page = browser.new_page(viewport={"width": 1920, "height": 1080})
        
        for file in html_files:
            file_path = os.path.join(base_dir, file)
            if not os.path.exists(file_path):
                print(f"File not found: {file_path}")
                continue
                
            print(f"Processing {file}...")
            # Load the local HTML file
            page_uri = "file:///" + file_path.replace("\\", "/")
            page.goto(page_uri, timeout=60000)
            
            # Wait for Chart.js animations to complete
            page.wait_for_timeout(2500)
            
            # Create a subfolder for this dashboard
            dashboard_name = file.replace(".html", "")
            dashboard_dir = os.path.join(output_dir, dashboard_name)
            if not os.path.exists(dashboard_dir):
                os.makedirs(dashboard_dir)
            
            # Take a full page screenshot
            print(f"  -> Saving full page screenshot...")
            page.screenshot(path=os.path.join(dashboard_dir, "Dashboard_Complet.png"), full_page=True)
            
            # Screenshot all canvas elements (the diagrams)
            canvases = page.locator("canvas").all()
            for i, canvas in enumerate(canvases):
                try:
                    name = canvas.get_attribute("id") or f"diag_{i+1}"
                    print(f"  -> Saving diagram {name}...")
                    canvas.screenshot(path=os.path.join(dashboard_dir, f"Diagramme_{name}.png"), animations="disabled")
                except Exception as e:
                    print(f"     Could not screenshot canvas {i}: {e}")
                    
            # Screenshot all glass-card elements (KPIs and sections)
            cards = page.locator(".glass-card").all()
            for i, card in enumerate(cards):
                try:
                    # check if card is actually visible
                    if card.is_visible():
                        print(f"  -> Saving KPI/Section card {i+1}...")
                        card.screenshot(path=os.path.join(dashboard_dir, f"KPI_Section_{i+1}.png"), animations="disabled")
                except Exception as e:
                    print(f"     Could not screenshot card {i}: {e}")
                    
        browser.close()
        print(f"Done! Images saved in {output_dir}")

if __name__ == "__main__":
    export_dashboards()

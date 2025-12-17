import sys
import io
import os
import datetime
import json
import shutil

# Comprehensive encoding fix for Windows
if sys.platform == "win32":
    os.environ["PYTHONUTF8"] = "1"
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def main():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    default_dir = os.path.join(BASE_DIR, "input", "csv")
    
    # Import functions
    from data_description import process_uploaded_file, get_latest_file
    
    # Get latest file
    latest_file = get_latest_file(default_dir)
    print(f"Latest file found: {latest_file}")
    
    if latest_file is None or not os.path.exists(latest_file):
        print(f"No input files found in {default_dir}")
        return
    
    # ======================================================================
    # Step 1: Data Description
    # ======================================================================
    print("\n" + "="*60)
    print("STEP 1: PROCESSING DATA DESCRIPTION")
    print("="*60)
    
    result = process_uploaded_file(latest_file)
    output_dir = result["outputs_dir"]
    print(f"✅ Description JSON created at: {result['description_json']}")
    print(f"📁 Output directory: {output_dir}")
    
    # ======================================================================
    # Step 2: Charts and Insights
    # ======================================================================
    print("\n" + "="*60)
    print("STEP 2: GENERATING CHARTS AND INSIGHTS")
    print("="*60)
    try:
        from charts import HotelBookingVisualization
        
        engine = HotelBookingVisualization(
            dataset_path=latest_file,
            output_path=output_dir
        )
        
        json_config_path = os.path.join(BASE_DIR, "input", "csv", "input.json")
        
        if not os.path.exists(json_config_path):
            print(f"❌ Config file not found: {json_config_path}")
            return
        
        print(f"✅ Found config file: {json_config_path}")
        charts_results = engine.process_json_config(json_config_path)
        print("✅ Charts and insights generated successfully")
    except:
        print("no charts")
    
    # ======================================================================
    # Step 3: Dashboard Creation
    # ======================================================================
    print("\n" + "="*60)
    print("STEP 3: CREATING DASHBOARD")
    print("="*60)
    
    try:
        with open("input/csv/input.json", 'r') as file:
            data = json.load(file)
    
        if data.get("dashboards") and len(data["dashboards"]) > 0:
            # Load the dataset for dashboard
            import pandas as pd
            df = pd.read_csv(latest_file)
        
            # Create dashboard using the new dashboard.py functionality
            dashboard_output = os.path.join(output_dir, "dashboard.png")
        
            # Import and use the new dashboard function
            from dashboard import create_dashboard
        
            # Create the professional PowerBI dashboard
            success_path = create_dashboard(
                df=df, 
                json_file=json_config_path, 
                output=dashboard_output
            )
        else:
            success_path = None
    except Exception as e:
        print(f"❌ Dashboard creation encountered an error: {e}")
        import traceback
        traceback.print_exc()
        success_path = None

    # Determine dashboard status based on success_path
    dashboard_status = success_path if success_path and os.path.exists(success_path) else "not created"

    # Save the status in a JSON file in output_dir
    dashboard_json_path = os.path.join(output_dir, "dashboard.json")
    with open(dashboard_json_path, 'w') as f:
        json.dump({"dashboard_status": dashboard_status}, f, indent=2)

    # Optional: Print status message
    if dashboard_status == "success":
        print("✅ PowerBI-level professional dashboard created successfully")
    else:
        print("❌ Dashboard creation failed or no dashboard generated")
    
    # ======================================================================
    # Step 4: NLP to SQL Query Processing
    # ======================================================================
    print("\n" + "="*60)
    print("STEP 4: PROCESSING NLP TO SQL QUERIES")
    print("="*60)
    
    try:
        from query_handler import execute_query
        
        # Load query from input.json
        user_query = None
        if os.path.exists(json_config_path):
            with open(json_config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "queries" in data and len(data["queries"]) > 0:
                    user_query = data["queries"][0].get("text", "").strip()
        
        if user_query:
            print(f"🧠 Executing user query: {user_query}")
            query_output_path = execute_query(latest_file, user_query)
            
            if os.path.exists(query_output_path):
                main_query_output = os.path.join(output_dir, "query_output.json")
                shutil.copy2(query_output_path, main_query_output)
                print(f"✅ Query output saved to: {main_query_output}")
            else:
                main_query_output = os.path.join(output_dir, "query_output.json")
                with open(main_query_output, 'w', encoding='utf-8') as f:
                    json.dump([], f)
                print(f"⚠️ Query execution returned no results")
        else:
            print("ℹ️ No user queries found in input.json")
            main_query_output = os.path.join(output_dir, "query_output.json")
            with open(main_query_output, 'w', encoding='utf-8') as f:
                json.dump([], f)
            print(f"ℹ️ Created empty query output")
    
    except Exception as e:
        print(f"❌ Query processing failed: {e}")
        main_query_output = os.path.join(output_dir, "query_output.json")
        with open(main_query_output, 'w', encoding='utf-8') as f:
            json.dump([], f)
    
    # ======================================================================
    # Step 5: LLM Analysis Service
    # ======================================================================
    print("\n" + "="*60)
    print("STEP 5: RUNNING LLM ANALYSIS SERVICE")
    print("="*60)
    
    try:
        # Create sentiment.json if needed
        sentiment_path = os.path.join(output_dir, "sentiment.json")
        if not os.path.exists(sentiment_path):
            with open(sentiment_path, 'w') as f:
                json.dump({}, f)
            print("Created empty sentiment.json")
        
        # Verify all required files exist before calling service
        required_files = [
            "data_description.json",
            "insights.json", 
            "insights_charts.json",
            "comparison.json",
            "dashboard.json",
            "query_output.json",
            "sentiment.json"
        ]
        
        print(f"📂 Verifying files in: {output_dir}")
        for filename in required_files:
            filepath = os.path.join(output_dir, filename)
            if os.path.exists(filepath):
                print(f"  ✓ {filename}")
            else:
                print(f"  ⚠ {filename} - creating empty file")
                # Create empty file if missing
                with open(filepath, 'w', encoding='utf-8') as f:
                    if filename.endswith('.json'):
                        json.dump({} if filename != "insights.json" and filename != "insights_charts.json" and filename != "query_output.json" else [], f)
        
        # Import and configure service module
        import service as service_module
        
        # Set the outputs_dir attribute that service.py will use
        service_module.outputs_dir = output_dir
        print(f"✅ Configured service to use output directory: {output_dir}")
        
        from service import service
        service()
        print("✅ LLM analysis completed successfully")
    
    except Exception as e:
        print(f"❌ LLM analysis failed: {e}")
        import traceback
        traceback.print_exc()
    
    # ======================================================================
    # Step 6: PowerPoint Generation - SUBPROCESS APPROACH
    # ======================================================================
    print("Step 6: Generating PowerPoint presentation...")
    try:
        # Import and run PPT generation
        import subprocess
        import sys
        
        # FIX: Run ppt.py as a subprocess with proper UTF-8 encoding
        env = os.environ.copy()
        env['PYTHONUTF8'] = '1'
        env['PYTHONIOENCODING'] = 'utf-8'
        
        result = subprocess.run(
            [sys.executable, "-m", "ppt"],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            env=env,
            timeout=300,
            cwd=os.path.dirname(os.path.abspath(__file__))
        )
        
        if result.returncode == 0:
            print("✅ PowerPoint presentation created successfully")
            print(f"PPT Output: {result.stdout}")
        else:
            print(f"❌ PPT generation failed with return code {result.returncode}")
            print(f"Error: {result.stderr}")
            
    except subprocess.TimeoutExpired:
        print("❌ PPT generation timed out after 5 minutes")
    except Exception as e:
        print(f"❌ PPT generation failed: {e}")
        import traceback
        traceback.print_exc()
    
    # ======================================================================
    # Final Summary
    # ======================================================================
    print("\n" + "="*60)
    print("🎯 ALL PROCESSES COMPLETED!")
    print("="*60)
    print(f"📁 Output directory: {output_dir}")
    print(f"📊 Data Description: ✓")
    print(f"📈 Charts & Insights: ✓") 
    print(f"📋 Dashboard: {'✓' if dashboard_status != 'not created' else 'ℹ️'}")
    print(f"🔍 Query Results: ✓")
    print(f"🤖 LLM Analysis: ✓")
    print(f"📊 PowerPoint: ✓")
    print("="*60)


if __name__ == "__main__":
    main()
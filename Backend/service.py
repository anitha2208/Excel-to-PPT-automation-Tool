import os
import json
import google.generativeai as genai

def load_json_if_exists(path, default={}):
    """Load JSON file if it exists - with better error handling"""
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"⚠️ Error loading {path}: {e}")
    return default
# ==============================
# Gemini Model Loader
# ==============================
def load_model():
    api_key = "AIzaSyCA-rO_ZqfVGGF1sO5BKVEGClxSA2UTezY"  # ⚡ Replace with your Gemini API key
    if not api_key or api_key.strip() == "":
        raise ValueError("❌ GEMINI_API_KEY is missing.")
    
    genai.configure(api_key=api_key)
    
    # List available models to debug
    print("🔍 Available models:")
    try:
        for model in genai.list_models():
            if 'generateContent' in model.supported_generation_methods:
                print(f"  - {model.name}")
    except Exception as e:
        print(f"⚠️ Could not list models: {e}")
    
    # Try different model names based on available models (prioritizing newer, faster models)
    model_names = [
        "models/gemini-2.5-flash",
        "models/gemini-2.0-flash",
        "models/gemini-2.5-pro",
        "models/gemini-2.0-pro-exp",
        "models/gemini-flash-latest",
        "models/gemini-pro-latest",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-flash-latest",
        "gemini-pro-latest"
    ]
    
    for model_name in model_names:
        try:
            print(f"🔄 Trying model: {model_name}")
            model = genai.GenerativeModel(model_name)
            print(f"✅ Successfully loaded model: {model_name}")
            return model
        except Exception as e:
            print(f"❌ Failed to load {model_name}: {e}")
            continue
    
    raise ValueError("❌ No compatible Gemini model found. Please check your API key and available models.")


# ==============================
# Build Prompt for LLM
# ==============================
def build_prompt(description, insights, comparison, dashboard, query_out, sentiment_data,insights_charts):
    json_schema = '''
{
  "slides": [
    {
      "slide_index": "integer",
      "placeholders": {
        "title": "string",
        "subtitle": "string (optional)",
        "content": "string (optional)",
        "image_path": "string (optional)"
      }
    }
  ]
}
'''

    prompt = f"""
You are an expert PPT creator and data analyst.

### Dataset Description
{json.dumps(description, indent=2)}

### Insights
{json.dumps(insights, indent=2)}

### Comparisons
{json.dumps(comparison, indent=2)}

### Dashboard
{json.dumps(dashboard, indent=2)}

### Query Results
{json.dumps(query_out, indent=2)}

### Sentiment Analysis
{json.dumps(sentiment_data, indent=2)}

### Insight_charts
{json.dumps(insights_charts, indent=2)}


---

Generate a structured JSON (preview.json) **strictly following this schema: {json_schema}**

** UNDERSTAND THIS YOUR MAKING PPT CONTENT GENERATION FOR BUSSINESS AND COMPANIES WHICH ARE PROFFESSIONAL , SO MAKE SURE TO HAVE CONTENT FOR THAT**
Rules:
- Slide 2: Data Description
- Slide 3: Key Comparisons
- Slide 4...N: Feature Insights (with plots if available)
- Next Slide: Query Results (Optional)
- Next Slide: Dashboard Highlights (Optional)
- Next slide : Insights_charts (optional) -- Check the json , if not empty -- Give Insights
- If sentiment.json is not empty → create a "Sentiment Analysis" slide
- Final Slides: Summary, Business Insights, Conclusion, Thank You

### CONTENT RULES
1. **Data Description (Slide 2)** → Convert dataset metadata (rows, columns, target, description) into 2–3 professional sentences.
Also the description to be more proffessional and give in a proper paragraph content that fits one slide in the PPT.
2. **Dashboard (Slide 3, if available)** → Show as one image slide, use provided path.
3. **Insights** → For each column, create 3–4 bullet points from statistics (mean, median, mode, min, max, etc.).Make sure to use the statistical data and create the points so proffessional a data analyst
4. **Charts + Insights** → Add both insights + chart path in slide.This should have a prope rinsight of data and not the same meta data or statistical data given to you , make it a proffessional analysis insights 
Make sure to use the statistical data and create the points so proffessional a data analyst
5. **Comparison** → Explain relationship (col1 vs col2) and use given chart path.
6. **Query Table** → Show sample table (10–15 rows), add row count if truncated.
7. **Business Insights** → 4 actionable business recommendations.
This is the most important part of the slide , try give a overall research anf insight for bussiness improvement **USE THE ENTIRE DATA GIVEN AND DO THIS ANALYSIS**.
8. **Summary** → Detailed recap of dataset trends. Make sure to cover everything and conclude well , give a paragraph of content that cover a slide in ppt
9. **Thank You** → Closing slide.Professional and good closing for the PPT

### OUTPUT Requirements: 
** If there is a path of any image or file , use the same in the preview.json where ever needed but that path need not to be displayed in the slide content (Content should be used for Bussiness presentation so make it proffessional)
# - Output only valid JSON (no markdown or extra text). 
# - Human readable , bussiness-focused content. 
# - Ensure all text is professional, business-focused, and concise. 
# - Position text and images with non-overlapping. 
# - Keep content human-readable and actionable.
Only output valid JSON (no markdown, no explanation).
"""
    return prompt.strip()


# ==============================
# Main Service Orchestrator
# ==============================
def service():
    outputs_dir = getattr(service, 'output_dir', 'output')
    print(outputs_dir)
    os.makedirs(outputs_dir, exist_ok=True)

    print(f"🔍 Looking for files in: {outputs_dir}")
    # outputs_dir = "outputs"
    # os.makedirs(outputs_dir, exist_ok=True)

    # # === Load all required files from outputs folder ===
    # def load_json_if_exists(path, default={}):
    #     try:
    #         if os.path.exists(path):
    #             with open(path, "r", encoding="utf-8") as f:
    #                 return json.load(f)
    #     except Exception as e:
    #         print(f"⚠️ Error loading {path}: {e}")
    #     return default

    description = load_json_if_exists(os.path.join(outputs_dir, "data_description.json"))
    insights_data = load_json_if_exists(os.path.join(outputs_dir, "insights.json"), [])
    charts_insights_data = load_json_if_exists(os.path.join(outputs_dir, "insights_charts.json"), [])
    comparison = load_json_if_exists(os.path.join(outputs_dir, "comparison.json"))
    dashboard_data = load_json_if_exists(os.path.join(outputs_dir, "dashboard.json"))
    query_out = load_json_if_exists(os.path.join(outputs_dir, "query_output.json"))
    sentiment_data = load_json_if_exists(os.path.join(outputs_dir, "sentiment.json"))

    # === Sanity check: at least description.json must exist ===
    if not description:
        raise FileNotFoundError(f"❌ description.json not found in {outputs_dir}. Required for LLM input.")

    # === LLM Processing ===
    print("⚡ Loading Gemini model...")
    try:
        model = load_model()
    except Exception as e:
        print(f"❌ Failed to load model: {e}")
        return

    print("📝 Building prompt...")
    prompt = build_prompt(description, insights_data,charts_insights_data, comparison, dashboard_data, query_out, sentiment_data)

    print("🚀 Sending request to Gemini...")
    try:
        # Add generation config for better control
        generation_config = {
            "temperature": 0.1,
            "top_p": 0.8,
            "top_k": 40,
            "max_output_tokens": 8192,
        }
        
        response = model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings=[
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            ]
        )
        
        generated_text = response.text.strip() if response and response.text else ""
        
    except Exception as e:
        print(f"❌ Error generating content: {e}")
        if "quota" in str(e).lower():
            print("💡 Quota exceeded. Try again later or check your API usage.")
        elif "api key" in str(e).lower():
            print("💡 API key issue. Please verify your Gemini API key.")
        return

    if not generated_text:
        print("❌ Empty response from Gemini")
        return

    # ✅ Clean if wrapped with ```
    if generated_text.startswith("```"):
        generated_text = generated_text.strip("`")
        if generated_text.lower().startswith("json"):
            generated_text = generated_text[4:].strip()

    try:
        preview_json = json.loads(generated_text)
        print("✅ Successfully parsed JSON from Gemini.")
    except json.JSONDecodeError as e:
        print(f"⚠️ Output was not valid JSON: {e}")
        print(f"Raw output: {generated_text[:500]}...")
        preview_json = {"raw_output": generated_text, "error": "Invalid JSON format"}

    out_path = os.path.join(outputs_dir, "preview.json")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(preview_json, f, indent=2, ensure_ascii=False)
        print(f"🎉 Preview JSON saved to {out_path}")
    except Exception as e:
        print(f"❌ Error saving file: {e}")


# ==============================
# Run Example
# ==============================
if __name__ == "__main__":
    service()
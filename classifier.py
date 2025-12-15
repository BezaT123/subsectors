import json
import pandas as pd
import openai
from typing import Dict, List, Optional, Tuple, Any
import os
import argparse
from dataclasses import dataclass

from extract import extract_setup_data_to_json

@dataclass
class BusinessSummary:
    business_name: str
    industry_type: str
    primary_industry: str
    secondary_industry: Optional[str]
    revenue_streams: List[str]
    operating_costs: List[str]
    products: List[str]
    cos_categories: List[str]  # Cost of Sales categories
    top_products: List[str]    # Top products (weighting-based)
    currency: str
    benchmark_metrics: Dict[str, Any]

@dataclass
class ClassificationResult:
    sector: str
    primary_subsector: str
    additional_subsectors: List[str]
    top_products: List[str]
    confidence_explanation: str

class BusinessClassifier:
    def __init__(self, openai_api_key: str, reference_file_path: str):
        """
        Initialize the Business Classifier
        
        Args:
            openai_api_key: OpenAI API key
            reference_file_path: Path to the SubSectors reference Excel file
        """
        self.client = openai.OpenAI(api_key=openai_api_key)
        self.reference_data = self._load_reference_data(reference_file_path)
    
    def _load_reference_data(self, file_path: str) -> Dict[str, List[str]]:
        """Load and parse the reference sectors and sub-sectors data"""
        try:
            df = pd.read_excel(file_path)
            # Assuming columns are: Business Sector | Example Sub-Sectors
            sectors_dict = {}
            
            for _, row in df.iterrows():
                sector = str(row.iloc[0]).strip()
                subsector = str(row.iloc[1]).strip()
                
                if sector not in sectors_dict:
                    sectors_dict[sector] = []
                
                if subsector and subsector != 'nan':
                    sectors_dict[sector].append(subsector)
            
            return sectors_dict
        except Exception as e:
            raise Exception(f"Error loading reference file: {str(e)}")
    
    def extract_business_summary(self, json_data: Dict) -> BusinessSummary:
        """Extract relevant business information from the JSON structure"""
        
        setup_data = json_data.get('i_Setup', {}).get('fields', {})
        cos_data = json_data.get('i_COS', {})
        info_metrics = json_data.get('info', {})
        
        # Extract basic business info
        business_name = setup_data.get('Business Name', {}).get('value', '')
        currency = setup_data.get('Currency', {}).get('value', '')
        
        # Extract industry details
        industry_details = setup_data.get('Industry Details', {}).get('subTableData', [])
        industry_type = ''
        primary_industry = ''
        secondary_industry = None
        
        for item in industry_details:
            if item.get('fieldLabel') == 'Industry Type':
                industry_type = item.get('value', '')
            elif item.get('fieldLabel') == 'Primary Industry':
                primary_industry = item.get('value', '')
            elif item.get('fieldLabel') == 'Secondary Industry':
                secondary_industry = item.get('value', '')
        
        # Extract revenue streams
        revenue_streams = []
        revenue_data = setup_data.get('Revenue Streams', {}).get('subTableData', [])
        for item in revenue_data:
            if item.get('name'):
                revenue_streams.append(item.get('name'))
        
        # Extract operating costs
        operating_costs = []
        opex_data = setup_data.get('Operating Costs', {}).get('subTableData', [])
        for item in opex_data:
            if item.get('name'):
                operating_costs.append(item.get('name'))
        
        # Extract products, Cost of Sales categories, and top products (by weighting only)
        products: List[str] = []
        cos_categories: List[str] = []
        top_products_section = cos_data.get('topProducts', {})
        top_products_weighting = top_products_section.get('byWeighting', []) if isinstance(top_products_section, dict) else []
        top_products_list: List[str] = []
        
        # Collect top products names from weighting list
        for item in top_products_weighting:
            name = str(item.get('productName', '')).strip()
            if name:
                top_products_list.append(name)
        
        # Fallback: derive from products list if needed
        cos_entries: List[Dict] = top_products_weighting if top_products_weighting else cos_data.get('products', [])
        
        for product in cos_entries:
            product_name = str(product.get('productName', '')).strip()
            cos_category = str(product.get('costOfSalesCategory', '')).strip()
            
            if product_name and product_name not in products:
                products.append(product_name)
            if cos_category and cos_category not in cos_categories:
                cos_categories.append(cos_category)
        
        # If we still lack top products, use first 5 product names as a fallback
        if not top_products_list:
            top_products_list = products[:5]
        
        return BusinessSummary(
            business_name=business_name,
            industry_type=industry_type,
            primary_industry=primary_industry,
            secondary_industry=secondary_industry,
            revenue_streams=revenue_streams,
            operating_costs=operating_costs,
            products=products,
            cos_categories=cos_categories,
            top_products=top_products_list,
            currency=currency,
            benchmark_metrics=info_metrics
        )
    
    def get_relevant_sectors(self, summary: BusinessSummary) -> Dict[str, List[str]]:
        """Filter reference data to get relevant sectors based on industry classification"""
        
        industries_to_match = [summary.primary_industry]
        if summary.industry_type.lower() == 'combined' and summary.secondary_industry:
            industries_to_match.append(summary.secondary_industry)
        
        relevant_sectors = {}
        
        # Direct matching first
        for industry in industries_to_match:
            for sector, subsectors in self.reference_data.items():
                if industry.lower().replace('/', ' ').strip() in sector.lower().replace('/', ' '):
                    relevant_sectors[sector] = subsectors
        
        # Fuzzy matching for common terms
        fuzzy_matches = {
            'trade': ['Retail / Trade'],
            'retail': ['Retail / Trade'],
            'manufacturing': ['Manufacturing'],
            'agriculture': ['Agri primary production'],
            'transport': ['Transport / logistics'],
            'healthcare': ['Healthcare'],
            'education': ['Education'],
            'hospitality': ['Hospitality']
        }
        
        for industry in industries_to_match:
            industry_lower = industry.lower()
            for term, sectors in fuzzy_matches.items():
                if term in industry_lower:
                    for sector_pattern in sectors:
                        matching_sectors = [s for s in self.reference_data.keys() 
                                          if sector_pattern.lower() in s.lower()]
                        for sector in matching_sectors:
                            relevant_sectors[sector] = self.reference_data[sector]
        
        return relevant_sectors
    
    def create_llm_prompt(self, summary: BusinessSummary, relevant_sectors: Dict[str, List[str]]) -> str:
        """Create a structured prompt for the LLM to classify the business"""
        
        prompt = f"""<ROLE> \n You are a business classification expert. Based on the business information provided, classify this business into the most appropriate sector and sub-sectors. Focus on sector and sub-sector selection; ignore level 3 codes. Use top products as contextual evidence.

BUSINESS INFORMATION:
- Business Name: {summary.business_name}
- Primary Industry: {summary.primary_industry}
- Secondary Industry: {summary.secondary_industry or 'None'}

REVENUE STREAMS:
{', '.join(summary.revenue_streams) if summary.revenue_streams else 'None specified'}

PRODUCTS/INVENTORY:
{', '.join(summary.products[:20]) if summary.products else 'None specified'}

TOP PRODUCTS (weighting-based):
{', '.join(summary.top_products[:5]) if summary.top_products else 'None found'}

COST OF SALES CATEGORIES:
{', '.join(summary.cos_categories) if summary.cos_categories else 'None specified'}

---
**BENCHMARKABLE OPERATIONAL METRICS (CONTEXT ONLY):**
{json.dumps(summary.benchmark_metrics, indent=4) if summary.benchmark_metrics else 'None found.'}
---

AVAILABLE SECTORS AND SUB-SECTORS TO CHOOSE FROM:
"""
        
        for sector, subsectors in relevant_sectors.items():
            prompt += f"\n{sector}:\n"
            for subsector in subsectors:
                prompt += f"  - {subsector}\n"
        
        prompt += """
CLASSIFICATION REQUIREMENTS:
1. Choose the most appropriate PRIMARY SECTOR from the list above
2. Select the best PRIMARY SUB-SECTOR from that sector
3. If applicable, identify any ADDITIONAL SUB-SECTORS (from the same or different sectors)
4. Do NOT return level 3 codes; instead, include the top products list (weighting-based) as part of the response.

RESPONSE FORMAT (JSON):
{
    "sector": "Selected primary sector",
    "primary_subsector": "Selected primary sub-sector",
    "additional_subsectors": ["Any additional relevant sub-sectors"],
    "top_products": ["Up to 5 top products (weighting-based)"],
    "confidence_explanation": "Brief explanation of why these classifications were chosen based on the business data and top products."
}

Provide only the JSON response, no additional text."""        
        return prompt
    
    def classify_with_llm(self, prompt: str) -> ClassificationResult:
        """Send prompt to OpenAI and parse the response"""
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a business classification expert. Respond only with valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=500
            )
            
            result_text = response.choices[0].message.content.strip()
            
            # Parse JSON response
            result_json = json.loads(result_text)
            
            return ClassificationResult(
                sector=result_json.get('sector', ''),
                primary_subsector=result_json.get('primary_subsector', ''),
                additional_subsectors=result_json.get('additional_subsectors', []),
                top_products=result_json.get('top_products', []),
                confidence_explanation=result_json.get('confidence_explanation', '')
            )
            
        except Exception as e:
            raise Exception(f"Error in LLM classification: {str(e)}")
    
    def classify_business(self, json_data: Dict) -> ClassificationResult:
        """Main method to classify a business from JSON data"""
        
        # Extract business summary
        summary = self.extract_business_summary(json_data)
        
        # Get relevant sectors
        relevant_sectors = self.get_relevant_sectors(summary)
        
        if not relevant_sectors:
            raise Exception(f"No matching sectors found for industries: {summary.primary_industry}, {summary.secondary_industry}")
        
        # Create LLM prompt
        prompt = self.create_llm_prompt(summary, relevant_sectors)
        
        # Get classification from LLM
        result = self.classify_with_llm(prompt)
        
        # Attach top products from summary into result if LLM omitted them
        if not result.top_products:
            result.top_products = summary.top_products[:5]
        
        return result

def main():
    """Example usage of the Business Classifier"""
    
    parser = argparse.ArgumentParser(description="Classify a business from Excel or extracted JSON.")
    parser.add_argument(
        "--file",
        "-f",
        required=True,
        help="Path to company Excel file (.xlsm/.xlsx) or extracted JSON (.json)",
    )
    args = parser.parse_args()

    # Configuration
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")  # Set this environment variable
    script_dir = os.path.dirname(os.path.abspath(__file__))
    REFERENCE_FILE_PATH = os.path.join(script_dir, "categories ideas.xlsx")  # Hardcoded reference
    
    if not OPENAI_API_KEY:
        raise ValueError("Please set OPENAI_API_KEY environment variable")
    
    # Initialize classifier
    classifier = BusinessClassifier(OPENAI_API_KEY, REFERENCE_FILE_PATH)
    
    # Load data: Excel -> run extractor; JSON -> load directly
    json_data = None
    file_lower = args.file.lower()
    if file_lower.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        json_data = extract_setup_data_to_json(args.file, output_json_path=None)
    elif file_lower.endswith(".json"):
        with open(args.file, "r", encoding="utf-8") as f:
            json_data = json.load(f)
    else:
        print("Unsupported file type. Use .xlsx/.xlsm or extracted .json.")
        return None

    if not json_data:
        print(f"Could not extract/load data from {args.file}")
        return None
    
    # Classify the business
    try:
        result = classifier.classify_business(json_data)
        
        print("BUSINESS CLASSIFICATION RESULTS:")
        print("=" * 40)
        print(f"Sector: {result.sector}")
        print(f"Primary Sub-sector: {result.primary_subsector}")
        print(f"Additional Sub-sectors: {', '.join(result.additional_subsectors) if result.additional_subsectors else 'None'}")
        print(f"Top Products: {', '.join(result.top_products[:5]) if result.top_products else 'None'}")
        print(f"\nConfidence Explanation: {result.confidence_explanation}")
        
        # Return structured result for further processing
        return {
            'sector': result.sector,
            'primary_subsector': result.primary_subsector,
            'additional_subsectors': result.additional_subsectors,
            'top_products': result.top_products[:5],
            'confidence_explanation': result.confidence_explanation
        }
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

if __name__ == "__main__":
    main()
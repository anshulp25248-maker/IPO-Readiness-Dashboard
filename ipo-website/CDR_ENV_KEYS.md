# CDR AI Keys

Use these six Groq keys and six Tavily keys for the CDR section. The app falls back to `GROQ_API_KEY` and `TAVILY_API_KEY` only when a section-specific key is blank.

The individual CDR tabs use their matching section keys. The Comprehensive CDR / DOCX flow runs the five section lanes first, so Sector Analysis uses the sector keys, Director Profile uses the director keys, Competitor Analysis uses the competitor keys, and so on. The comprehensive key is used only to synthesize the final recommendation after those section reports are generated.

```env
# CDR Section 1: Sector Analysis
CDR_SECTOR_ANALYSIS_GROQ_API_KEY=
CDR_SECTOR_ANALYSIS_TAVILY_API_KEY=

# CDR Section 2: Industry Analysis
CDR_INDUSTRY_ANALYSIS_GROQ_API_KEY=
CDR_INDUSTRY_ANALYSIS_TAVILY_API_KEY=

# CDR Section 3: Competitor Analysis
CDR_COMPETITOR_ANALYSIS_GROQ_API_KEY=
CDR_COMPETITOR_ANALYSIS_TAVILY_API_KEY=

# CDR Section 4: Director Profile
CDR_DIRECTOR_PROFILE_GROQ_API_KEY=
CDR_DIRECTOR_PROFILE_TAVILY_API_KEY=

# CDR Section 5: Company Analysis
CDR_COMPANY_ANALYSIS_GROQ_API_KEY=
CDR_COMPANY_ANALYSIS_TAVILY_API_KEY=

# CDR Section 6: Comprehensive CDR / DOCX Generation
CDR_COMPREHENSIVE_CDR_GROQ_API_KEY=
CDR_COMPREHENSIVE_CDR_TAVILY_API_KEY=
```

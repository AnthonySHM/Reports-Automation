# AI-Powered Cyber Insights

## Overview

The report generation system now includes AI-powered observations and recommendations on **Slide 3 (Cyber Insight Summary)**. These insights are automatically generated based on the report's security metrics to provide contextual analysis of the organization's security posture.

## Features

### Dynamic Insight Generation

The system analyzes available metrics and generates relevant observations:

- **Patch Management Analysis**: Evaluates the number of pending Microsoft KB patches and software updates
- **SOC Activity Assessment**: Contextualizes SOC ticket volumes and their implications
- **Detection Analysis**: Provides perspective on key security detection volumes
- **Network Vulnerability Review**: Comments on internal network vulnerabilities when available

### Intelligent Prioritization

The insight generator selects the most relevant observation based on:

1. **Data Significance**: Prioritizes metrics with notable values (high/low extremes)
2. **Contextual Relevance**: Chooses observations that provide actionable context
3. **Professional Tone**: Maintains appropriate language for security reporting

## Implementation

### Files Modified

#### `core/ai_insights.py` (New)
- `generate_cyber_insight()`: Main function that analyzes metrics and generates insights
- `generate_simple_insight()`: Fallback for when only patch data is available
- Configurable insight templates for various scenarios

#### `core/drive_agent.py`
- Added `replace_ai_insight_in_slide()` function to replace the placeholder text
- Searches for `[Additional observations and recommendations to be populated]` placeholder
- Replaces with AI-generated insight while preserving bullet formatting

#### `api.py`
- Imports `generate_cyber_insight` and `replace_ai_insight_in_slide`
- Generates insight after KB count replacement
- Currently uses available patch metrics (KB count, software count)
- Gracefully handles errors with fallback

### Template Change

#### `create_template.py`
The placeholder text on Slide 3 at line 439:
```python
"•  [Additional observations and recommendations to be populated]",
```

This is dynamically replaced during report generation.

## Example Insights

### High Patch Count
```
The environment shows a moderate patching backlog that should be addressed 
systematically to reduce the attack surface.
```

### Moderate Patch Count
```
The organization maintains a manageable patch status that should be addressed 
in the next maintenance window.
```

### Low Patch Count
```
The environment demonstrates strong patch management with minimal pending 
updates remaining.
```

### Excellent Status
```
The environment maintains up-to-date patch status with no pending critical 
updates identified during the reporting period.
```

### Formatting
- **No specific numbers**: Insights use qualitative descriptions instead of exact counts
- **Not bolded**: Text appears in regular weight to match other bullet points

## Future Enhancements

### Additional Metrics (Planned)
The system is designed to incorporate additional metrics when available:

- **Key Detections**: Number of security events detected
- **SOC Tickets**: Volume of SOC escalations
- **Analyst Investigations**: Depth of security review
- **Nuclei Vulnerabilities**: Internal network vulnerability count

### Enhancement Steps
1. Extract these metrics from the manifest or data sources
2. Pass them to `generate_cyber_insight()` in `api.py`
3. The insight generator will automatically prioritize the most relevant observation

## Testing

### Unit Tests
```bash
# Test insight generation with various scenarios
python core/ai_insights.py
```

### Integration Test
```bash
# Test complete replacement flow
python test_ai_insight.py
```

### Manual Verification
1. Generate a report via the API
2. Open the generated PowerPoint
3. Navigate to Slide 3 (Cyber Insight Summary)
4. Verify the last bullet point contains a contextual observation (not the placeholder)

## Error Handling

The system includes robust error handling:

- If insight generation fails, the placeholder text remains (manual review required)
- Errors are logged but don't block report generation
- Fallback to simple patch-based insights if advanced metrics unavailable

## Configuration

No configuration is required. The system automatically:

1. Detects available metrics from the report data
2. Generates appropriate insights
3. Replaces the placeholder on Slide 3

## Notes

- Insights are single-sentence observations for brevity
- Professional, objective tone suitable for client-facing reports
- Designed to complement (not replace) the detailed metrics in subsequent slides
- Works for both single-sensor and multi-sensor report configurations

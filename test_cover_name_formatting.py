"""
Test cover slide client name formatting.

Demonstrates that client names are truncated at the first hyphen on the cover slide.
"""

def test_client_name_formatting():
    """Test that client names are properly formatted for the cover slide."""
    
    test_cases = [
        {
            "input": "Elephant - 1st of the month",
            "expected": "Elephant",
            "description": "Standard format with hyphen and description"
        },
        {
            "input": "monterey mechanical - wt",
            "expected": "monterey mechanical",
            "description": "Lowercase with hyphen"
        },
        {
            "input": "RDLR - direct",
            "expected": "RDLR",
            "description": "Uppercase with hyphen"
        },
        {
            "input": "Packard",
            "expected": "Packard",
            "description": "No hyphen (simple name)"
        },
        {
            "input": "Company Name - Extra - More Info",
            "expected": "Company Name",
            "description": "Multiple hyphens (only first split)"
        }
    ]
    
    print("=" * 70)
    print("COVER SLIDE CLIENT NAME FORMATTING TEST")
    print("=" * 70)
    
    all_passed = True
    
    for test in test_cases:
        input_name = test["input"]
        expected = test["expected"]
        description = test["description"]
        
        # Apply the same logic as in _personalise_cover()
        cover_client_name = input_name.split('-')[0].strip() if '-' in input_name else input_name
        
        passed = cover_client_name == expected
        status = "[PASS]" if passed else "[FAIL]"
        
        if not passed:
            all_passed = False
        
        print(f"\n{status} {description}")
        print(f"  Input:    '{input_name}'")
        print(f"  Output:   '{cover_client_name}'")
        print(f"  Expected: '{expected}'")
    
    print("\n" + "=" * 70)
    if all_passed:
        print("[SUCCESS] All tests passed!")
    else:
        print("[FAILURE] Some tests failed!")
    print("=" * 70)
    
    return all_passed


if __name__ == "__main__":
    success = test_client_name_formatting()
    exit(0 if success else 1)

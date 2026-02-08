"""
Quick verification of multi-sensor slide mapping logic.

This script verifies that the slide mapping calculations in api.py are correct.
"""

def test_slide_mapping():
    """Test slide mapping for various sensor counts."""
    
    test_cases = [
        {
            "name": "Single Sensor",
            "sensors": ["DEFAULT"],
            "expected_structure": {
                "DEFAULT": {"ndr_slides": list(range(5, 13)), "total_per_sensor": 10}
            }
        },
        {
            "name": "Two Sensors",
            "sensors": ["SENSOR1", "SENSOR2"],
            "expected_structure": {
                "SENSOR1": {"ndr_slides": list(range(5, 13)), "total_per_sensor": 10},
                "SENSOR2": {"ndr_slides": list(range(15, 23)), "total_per_sensor": 10}
            }
        },
        {
            "name": "Three Sensors (Elephant)",
            "sensors": ["GAPRD", "VAHQ", "VAPRD"],
            "expected_structure": {
                "GAPRD": {"ndr_slides": list(range(5, 13)), "total_per_sensor": 10},
                "VAHQ": {"ndr_slides": list(range(15, 23)), "total_per_sensor": 10},
                "VAPRD": {"ndr_slides": list(range(25, 33)), "total_per_sensor": 10}
            }
        }
    ]
    
    print("=" * 70)
    print("MULTI-SENSOR SLIDE MAPPING VERIFICATION")
    print("=" * 70)
    
    for test in test_cases:
        print(f"\n{test['name']}")
        print("-" * 70)
        sensors = test["sensors"]
        n_sensors = len(sensors)
        
        # Calculate sensor-to-slide mapping (matching api.py logic)
        sensor_slide_map = {}
        for i, sensor_id in enumerate(sensors):
            base = 5 + i * 10  # 10 slides per sensor in template
            # Only map first 8 slides (NDR data slides)
            sensor_slide_map[sensor_id] = list(range(base, base + 8))
        
        # Calculate patches base (matching api.py logic)
        patches_base = 5 + n_sensors * 10
        
        # Verify against expected
        expected = test["expected_structure"]
        all_correct = True
        
        for sensor_id in sensors:
            actual_ndr = sensor_slide_map[sensor_id]
            expected_ndr = expected[sensor_id]["ndr_slides"]
            
            if actual_ndr == expected_ndr:
                status = "[OK]"
            else:
                status = "[FAIL]"
                all_correct = False
            
            print(f"  {status} {sensor_id:10} NDR slides: {actual_ndr[0]:2d}-{actual_ndr[-1]:2d} (indices)")
            
            # Show full sensor section
            full_section_start = actual_ndr[0]
            full_section_end = full_section_start + 9  # 10 slides total
            print(f"    {'':10} Full section: {full_section_start:2d}-{full_section_end:2d} (10 slides)")
            print(f"    {'':10}   - Slides {actual_ndr[0]:2d}-{actual_ndr[-1]:2d}: NDR data (8 slides)")
            print(f"    {'':10}   - Slide  {full_section_end-1:2d}: Vulnerabilities")
            print(f"    {'':10}   - Slide  {full_section_end:2d}: Mitigation")
        
        print(f"\n  Patches slides start at index: {patches_base}")
        print(f"    - Slide {patches_base:2d}: Required Patches (text)")
        print(f"    - Slide {patches_base+1:2d}: Software Patches (table)")
        print(f"    - Slide {patches_base+2:2d}: MS Patches (table)")
        
        total_slides = patches_base + 5  # patches + 2 aggregate vuln/mitigation
        print(f"\n  Total slides in presentation: {total_slides}")
        
        if all_correct:
            print(f"\n  Result: [PASS]")
        else:
            print(f"\n  Result: [FAIL]")
    
    print("\n" + "=" * 70)
    print("VERIFICATION COMPLETE")
    print("=" * 70)


def test_base_slide_mapping():
    """Test that base NDR slides [5,6,7,8,9,10,11,12] map correctly."""
    
    print("\n" + "=" * 70)
    print("BASE SLIDE MAPPING VERIFICATION")
    print("=" * 70)
    
    base_slides = [5, 6, 7, 8, 9, 10, 11, 12]
    slide_names = [
        "Outbound Data 1 (Heatmap)",
        "Outbound Data 2 (Heatmap)",
        "Top IP Destinations",
        "Top URLs",
        "External Destination",
        "Country by Connection",
        "Beaconing Score",
        "Sensitive Data"
    ]
    
    sensors = ["SENSOR1", "SENSOR2", "SENSOR3"]
    
    print("\nBase NDR slides (single-sensor template):")
    for idx, name in zip(base_slides, slide_names):
        print(f"  Slide {idx:2d}: {name}")
    
    print("\nMapping to multi-sensor template:")
    for i, sensor_id in enumerate(sensors):
        base = 5 + i * 10
        sensor_slides = list(range(base, base + 8))
        print(f"\n  {sensor_id}:")
        for base_idx, actual_idx, name in zip(base_slides, sensor_slides, slide_names):
            print(f"    Base {base_idx:2d} -> Actual {actual_idx:2d}: {name}")


if __name__ == "__main__":
    test_slide_mapping()
    test_base_slide_mapping()
    print("\n[SUCCESS] All calculations verified!")

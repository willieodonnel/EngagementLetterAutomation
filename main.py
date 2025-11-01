"""
Main Orchestration Module for Engagement Letter Automation

This module provides the main entry points and orchestration logic
for generating engagement letters. It coordinates between the data
processing and document generation modules.
"""

import os
from typing import Optional
from processing_data import (
    collect_engagement_data,
    collect_dual_engagement_data,
    save_data_to_json,
    load_data_from_json
)
from generating_doc import (
    generate_engagement_letter,
    generate_dual_engagement_letters,
    validate_template_exists,
    list_available_templates
)


def create_single_engagement_letter(
    loan_type: Optional[str] = None,
    letter_type: Optional[str] = None,
    use_autofill: bool = True,
    save_json: bool = False,
    database_path: str = "Templates/Approved Appraisers.csv",
    template_dir: str = "Templates",
    output_dir: str = "."
) -> str:
    """
    Create a single engagement letter with interactive or programmatic input.

    Args:
        loan_type: Type of loan ('7a', '504', 'CC'). If None, prompts user.
        letter_type: Type of letter ('App', 'Env', 'Sec', 'SFR', etc.). If None, prompts user.
        use_autofill: Whether to use autofill features (database lookup, date calculation)
        save_json: Whether to save the collected data to a JSON file
        database_path: Path to vendor database CSV file
        template_dir: Directory containing Word templates
        output_dir: Directory to save generated documents

    Returns:
        Path to the generated document
    """

    print("\n=== Single Engagement Letter Generation ===\n")

    # Get loan type if not provided
    if not loan_type:
        loan_type = input("Enter loan type (7a/504/CC): ").strip().upper()

    # Get letter type if not provided
    if not letter_type:
        letter_type = input("Enter letter type (App/Env/Sec/SFR/Phase 1/Phase 2): ").strip().upper()

    # Validate template exists
    if not validate_template_exists(loan_type, letter_type, template_dir):
        print(f"\nError: No template found for {loan_type} - {letter_type}")
        print(f"Available templates in {template_dir}:")
        for template in list_available_templates(template_dir):
            print(f"  - {template}")
        return None

    # Collect all necessary data
    print(f"\n--- Collecting Data for {loan_type} {letter_type} Letter ---\n")

    data = collect_engagement_data(
        loan_type=loan_type,
        letter_type=letter_type,
        use_autofill=use_autofill,
        database_path=database_path if use_autofill else None
    )

    # Save to JSON if requested
    if save_json:
        loan_name = data["loan"]["loan_name"]
        json_filename = f"{loan_name}_{letter_type}_data.json"
        save_data_to_json(data, json_filename)
        print(f"\nData saved to: {json_filename}")

    # Generate the document
    print("\n--- Generating Document ---\n")

    try:
        doc_path = generate_engagement_letter(
            data=data,
            template_dir=template_dir,
            output_dir=output_dir
        )
        print(f"\n✓ Successfully created: {doc_path}")
        return doc_path

    except Exception as e:
        print(f"\n✗ Error generating document: {e}")
        return None


def create_dual_engagement_letters(
    loan_type: Optional[str] = None,
    use_autofill: bool = True,
    save_json: bool = False,
    database_path: str = "Templates/Approved Appraisers.csv",
    template_dir: str = "Templates",
    output_dir: str = "."
) -> tuple:
    """
    Create both appraisal and environmental engagement letters.

    Args:
        loan_type: Type of loan ('7a', '504', 'CC'). If None, prompts user.
        use_autofill: Whether to use autofill features
        save_json: Whether to save the collected data to a JSON file
        database_path: Path to vendor database CSV file
        template_dir: Directory containing Word templates
        output_dir: Directory to save generated documents

    Returns:
        Tuple of paths to generated documents (appraisal_path, environmental_path)
    """

    print("\n=== Dual Engagement Letter Generation ===")
    print("(Creating both Appraisal and Environmental letters)\n")

    # Get loan type if not provided
    if not loan_type:
        loan_type = input("Enter loan type (7a/504/CC): ").strip().upper()

    # Validate templates exist
    if not validate_template_exists(loan_type, "APP", template_dir):
        print(f"\nError: Appraisal template not found for {loan_type}")
        return None, None

    if not validate_template_exists(loan_type, "ENV", template_dir):
        print(f"\nError: Environmental template not found for {loan_type}")
        return None, None

    # Collect all necessary data
    print(f"\n--- Collecting Data for {loan_type} Dual Letters ---\n")

    data = collect_dual_engagement_data(
        loan_type=loan_type,
        use_autofill=use_autofill,
        database_path=database_path if use_autofill else None
    )

    # Save to JSON if requested
    if save_json:
        loan_name = data["shared"]["loan"]["loan_name"]
        json_filename = f"{loan_name}_dual_data.json"
        save_data_to_json(data, json_filename)
        print(f"\nData saved to: {json_filename}")

    # Generate both documents
    print("\n--- Generating Documents ---\n")

    try:
        app_path, env_path = generate_dual_engagement_letters(
            dual_data=data,
            template_dir=template_dir,
            output_dir=output_dir
        )
        print(f"\n✓ Successfully created:")
        print(f"  - {app_path}")
        print(f"  - {env_path}")
        return app_path, env_path

    except Exception as e:
        print(f"\n✗ Error generating documents: {e}")
        return None, None


def create_from_json_file(
    json_path: str,
    template_dir: str = "Templates",
    output_dir: str = "."
) -> str:
    """
    Create engagement letter(s) from a saved JSON file.

    Args:
        json_path: Path to JSON file containing engagement data
        template_dir: Directory containing Word templates
        output_dir: Directory to save generated documents

    Returns:
        Path(s) to generated document(s)
    """

    print(f"\n=== Generating from JSON: {json_path} ===\n")

    try:
        data = load_data_from_json(json_path)

        # Check if it's dual data
        if "appraisal" in data and "environmental" in data:
            app_path, env_path = generate_dual_engagement_letters(
                dual_data=data,
                template_dir=template_dir,
                output_dir=output_dir
            )
            print(f"\n✓ Successfully created:")
            print(f"  - {app_path}")
            print(f"  - {env_path}")
            return f"{app_path}, {env_path}"
        else:
            doc_path = generate_engagement_letter(
                data=data,
                template_dir=template_dir,
                output_dir=output_dir
            )
            print(f"\n✓ Successfully created: {doc_path}")
            return doc_path

    except Exception as e:
        print(f"\n✗ Error: {e}")
        return None


def batch_create_from_json(
    json_dir: str = ".",
    template_dir: str = "Templates",
    output_dir: str = "output"
) -> list:
    """
    Batch process multiple JSON files to create engagement letters.

    Args:
        json_dir: Directory containing JSON files
        template_dir: Directory containing Word templates
        output_dir: Directory to save generated documents

    Returns:
        List of generated document paths
    """

    print(f"\n=== Batch Processing JSON Files from {json_dir} ===\n")

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    generated_docs = []
    json_files = [f for f in os.listdir(json_dir) if f.endswith('.json')]

    if not json_files:
        print(f"No JSON files found in {json_dir}")
        return generated_docs

    for json_file in json_files:
        print(f"\nProcessing: {json_file}")
        json_path = os.path.join(json_dir, json_file)

        result = create_from_json_file(json_path, template_dir, output_dir)
        if result:
            generated_docs.append(result)

    print(f"\n=== Batch Processing Complete ===")
    print(f"Generated {len(generated_docs)} document(s)")

    return generated_docs


def interactive_menu():
    """
    Interactive menu for choosing engagement letter generation options.
    """

    print("\n" + "=" * 60)
    print(" ENGAGEMENT LETTER AUTOMATION SYSTEM")
    print("=" * 60)

    while True:
        print("\nSelect an option:")
        print("1. Create single engagement letter")
        print("2. Create dual engagement letters (Appraisal + Environmental)")
        print("3. Generate from saved JSON file")
        print("4. Batch process JSON files")
        print("5. List available templates")
        print("6. Exit")

        choice = input("\nEnter choice (1-6): ").strip()

        if choice == "1":
            # Single letter generation
            use_autofill = input("\nUse autofill features? (y/n): ").strip().lower() == 'y'
            save_json = input("Save data to JSON? (y/n): ").strip().lower() == 'y'

            create_single_engagement_letter(
                use_autofill=use_autofill,
                save_json=save_json
            )

        elif choice == "2":
            # Dual letter generation
            use_autofill = input("\nUse autofill features? (y/n): ").strip().lower() == 'y'
            save_json = input("Save data to JSON? (y/n): ").strip().lower() == 'y'

            create_dual_engagement_letters(
                use_autofill=use_autofill,
                save_json=save_json
            )

        elif choice == "3":
            # Generate from JSON
            json_path = input("\nEnter JSON file path: ").strip()
            if os.path.exists(json_path):
                create_from_json_file(json_path)
            else:
                print(f"File not found: {json_path}")

        elif choice == "4":
            # Batch process
            json_dir = input("\nEnter directory containing JSON files (or press Enter for current): ").strip() or "."
            output_dir = input("Enter output directory (or press Enter for 'output'): ").strip() or "output"

            batch_create_from_json(json_dir, output_dir=output_dir)

        elif choice == "5":
            # List templates
            print("\nAvailable templates:")
            templates = list_available_templates()
            if templates:
                for template in templates:
                    print(f"  - {template}")
            else:
                print("  No templates found in Templates directory")

        elif choice == "6":
            # Exit
            print("\nGoodbye!")
            break

        else:
            print("\nInvalid choice. Please try again.")

        input("\nPress Enter to continue...")


# Programmatic API Examples

def example_programmatic_usage():
    """
    Examples of how to use the API programmatically without user interaction.
    """

    # Example 1: Create a single letter with all data provided
    data = {
        "loan_type": "7a",
        "letter_type": "App",
        "vendor": {
            "first_name": "Mark",
            "last_name": "Prottas",
            "company": "CBRE",
            "email": "mark.prottas@cbre.com",
            "type": "App",
            "region": "California"
        },
        "dates": {
            "current_date": "10/31/2024",
            "delivery_date": "11/14/2024"
        },
        "loan": {
            "loan_name": "ABC Property Development",
            "loan_number": "SBA-123456",
            "cdc_company": "N/A"
        },
        "property": {
            "property_address": "123 Main St, San Francisco, CA",
            "property_type": "Commercial Office",
            "sqft": "50000",
            "fee": "2500",
            "property_contact_name": "John Doe",
            "property_contact_phone": "(555) 123-4567",
            "item_to_send": "TBD"
        }
    }

    # Generate directly from data
    doc_path = generate_engagement_letter(data)
    print(f"Generated: {doc_path}")

    # Example 2: Create dual letters with partial data (will prompt for missing)
    from processing_data import collect_dual_engagement_data

    dual_data = collect_dual_engagement_data(
        loan_type="504",
        app_vendor_info={
            "first_name": "Mark",
            "last_name": "Prottas",
            "company": "CBRE",
            "email": "mark@cbre.com",
            "type": "App",
            "region": "CA"
        },
        env_vendor_info={
            "first_name": "Darrin",
            "last_name": "Domingo",
            "company": "Phase One Inc.",
            "email": "darrin@phaseone.com",
            "type": "Env",
            "region": "CA"
        },
        use_autofill=False  # Don't use database lookup
    )

    app_path, env_path = generate_dual_engagement_letters(dual_data)
    print(f"Generated: {app_path} and {env_path}")


if __name__ == "__main__":
    # Run the interactive menu when script is executed directly
    interactive_menu()
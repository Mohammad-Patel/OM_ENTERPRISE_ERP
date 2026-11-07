def amount_in_words(amount):
    """Convert numeric amount to Indian currency words format."""
    try:
        import num2words
        return num2words.num2words(amount, lang='en_IN').upper() + " ONLY"
    except ImportError:
        # Fallback if num2words library is not installed
        return f"RUPEES {amount} ONLY"
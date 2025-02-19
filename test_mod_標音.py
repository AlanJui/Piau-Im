from mod_標音 import split_tai_gi_im_piau


def test_split_tai_gi_im_piau():
    # test_cases = [
    #     ("but", ["b", "u", "4"]),
    #     ("bat", ["b", "a", "4"]),
    #     ("bak", ["b", "a", "4"]),
    #     ("bah", ["b", "a", "4"]),
    #     ("ka", ["k", "a", "1"]),
    #     ("ki", ["k", "i", "1"]),
    #     ("ku", ["k", "u", "1"]),
    #     ("pat", ["p", "a", "4"]),
    #     ("kak", ["k", "a", "4"]),
    #     ("tah", ["t", "a", "4"]),
    #     ("ho", ["h", "o", "1"]),
    # ]
    test_cases = [
        ("but", ["b", "ut", "4"]),
        ("bat", ["b", "at", "4"]),
        ("bak", ["b", "ak", "4"]),
        ("bah", ["b", "ah", "4"]),
        ("ka", ["k", "a", "1"]),
        ("ki", ["k", "i", "1"]),
        ("ku", ["k", "u", "1"]),
        ("pat", ["p", "at", "4"]),
        ("kak", ["k", "ak", "4"]),
        ("tah", ["t", "ah", "4"]),
        ("ho", ["h", "o", "1"]),
    ]

    for im_piau, expected in test_cases:
        result = split_tai_gi_im_piau(im_piau)
        assert result == expected, f"Test failed for {im_piau}: expected {expected}, got {result}"
        print(f"Test passed for {im_piau}: {result}")

if __name__ == "__main__":
    test_split_tai_gi_im_piau()

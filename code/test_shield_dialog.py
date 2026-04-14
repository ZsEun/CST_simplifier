"""Test the shield can classifier dialog standalone (no CST needed).

Run: python -m code.test_shield_dialog
"""

import tkinter as tk
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from code.shield_can_dialog import ShieldCanClassifierDialog, classify_shield_components

def test_classify():
    """Test keyword classification."""
    solids = [
        ("SUNRAY/SHIELD_COVER_ASM", "SUNRAY_PD_LED_SHIELDING_COVER_2"),
        ("SUNRAY/SHIELD_COVER_ASM", "SUNRAY_PD_LED_SHIELDING_COVER_3"),
        ("SUNRAY/SHIELD_FRAME_ASM", "SUNRAY_PD_LED_SHIELDING_FRAM_11"),
        ("SUNRAY/SHIELD_ASM", "SUNRAY_PD_SHIELDING_CAN_1"),
        ("SUNRAY/PCB_ASM", "SUNRAY_PD_LED_BOARD"),
    ]
    result = classify_shield_components(solids)
    print("Classification test:")
    print(f"  Covers: {[e['solid'] for e in result['cover']]}")
    print(f"  Frames: {[e['solid'] for e in result['frame']]}")
    print(f"  One piece: {[e['solid'] for e in result['one_piece']]}")
    assert len(result["cover"]) == 2, f"Expected 2 covers, got {len(result['cover'])}"
    assert len(result["frame"]) == 1, f"Expected 1 frame, got {len(result['frame'])}"
    assert len(result["one_piece"]) == 1, f"Expected 1 one_piece, got {len(result['one_piece'])}"
    print("  PASS")

def test_dialog():
    """Test the dialog interactively."""
    classified = {
        "cover": [
            {"comp": "SUNRAY/SHIELD_COVER_ASM", "solid": "SHIELDING_COVER_2", "shape": "SUNRAY/SHIELD_COVER_ASM:SHIELDING_COVER_2"},
            {"comp": "SUNRAY/SHIELD_COVER_ASM", "solid": "SHIELDING_COVER_3", "shape": "SUNRAY/SHIELD_COVER_ASM:SHIELDING_COVER_3"},
        ],
        "frame": [
            {"comp": "SUNRAY/SHIELD_FRAME_ASM", "solid": "SHIELDING_FRAM_11", "shape": "SUNRAY/SHIELD_FRAME_ASM:SHIELDING_FRAM_11"},
        ],
        "one_piece": [
            {"comp": "SUNRAY/SHIELD_ASM", "solid": "SHIELDING_CAN_1", "shape": "SUNRAY/SHIELD_ASM:SHIELDING_CAN_1"},
        ],
    }

    select_log = []
    def mock_cst_select(shape):
        select_log.append(shape)
        print(f"  [MOCK] SelectTreeItem: {shape}")

    root = tk.Tk()
    root.title("Test Host")
    root.geometry("200x100")

    def _open_dialog():
        dialog = ShieldCanClassifierDialog(root, classified, cst_select_fn=mock_cst_select)
        dialog.wait_window()
        print(f"\nDialog result: {dialog.result}")
        if dialog.result:
            for group in ["cover", "frame", "one_piece"]:
                items = dialog.result.get(group, [])
                print(f"  {group}: {[e['solid'] for e in items]}")
        root.quit()

    tk.Button(root, text="Open Dialog", command=_open_dialog).pack(pady=20)
    root.mainloop()

if __name__ == "__main__":
    print("=== Test 1: Classification ===")
    test_classify()
    print("\n=== Test 2: Dialog (interactive) ===")
    print("Instructions:")
    print("  1. Click 'Open Dialog'")
    print("  2. Select an item in a list, click 'Select in CST' — should print mock message")
    print("  3. Select an item, choose a target from dropdown — should move it")
    print("  4. Select an item, click 'Remove' — should remove it")
    print("  5. Type a name in 'Add component', pick a group, click 'Add'")
    print("  6. Click OK — should print the final lists")
    test_dialog()

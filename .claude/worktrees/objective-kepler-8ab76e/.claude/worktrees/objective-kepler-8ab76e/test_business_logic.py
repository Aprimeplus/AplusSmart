import unittest
# สมมติว่าไฟล์ business_logic.py อยู่ในโฟลเดอร์เดียวกัน
from business_logic import calculate_commission_data

class TestCommissionCalculations(unittest.TestCase):

    def test_plan_a_high_margin(self):
        """ทดสอบ Plan A กรณีที่ Margin >= 10%"""
        # เตรียมข้อมูลทดสอบ: ยอดขาย 1000, ต้นทุน 800 -> GP = 200, Margin = 20%
        mock_data = {"sales_service_amount": 1000}
        result = calculate_commission_data("Plan A", 1000, 800, mock_data)
        # ค่าคอมที่คาดหวัง: 35% ของ GP (200) = 70
        self.assertAlmostEqual(result["commission_amount"], 70)
        self.assertEqual(result["status"], "Normal")

    def test_plan_a_low_margin(self):
        """ทดสอบ Plan A กรณีที่ Margin < 10%"""
        # เตรียมข้อมูลทดสอบ: ยอดขาย 1000, ต้นทุน 950 -> GP = 50, Margin = 5%
        mock_data = {"sales_service_amount": 1000}
        result = calculate_commission_data("Plan A", 1000, 950, mock_data)
        # ค่าคอมที่คาดหวัง: 17.5% ของ GP (50) = 8.75
        self.assertAlmostEqual(result["commission_amount"], 8.75)
        self.assertEqual(result["status"], "Below T")

    def test_plan_b_not_eligible(self):
        """ทดสอบ Plan B กรณียอดขายไม่ถึงเกณฑ์"""
        mock_data = {"sales_service_amount": 499999}
        result = calculate_commission_data("Plan B", 0, 0, mock_data)
        self.assertAlmostEqual(result["commission_amount"], 0)
        self.assertEqual(result["status"], "Not Eligible (<500K)")

    def test_plan_b_full_tiers(self):
        """ทดสอบ Plan B กรณียอดขายครอบคลุมทุก Tier"""
        mock_data = {"sales_service_amount": 2500000} # ฐานคำนวณ = 2.4M
        result = calculate_commission_data("Plan B", 0, 0, mock_data)
        # Tier 1: 1,000,000 * 1.25% = 12,500
        # Tier 2: 1,000,000 * 1.75% = 17,500
        # Tier 3:   400,000 * 2.25% =  9,000
        # Total = 39,000
        self.assertAlmostEqual(result["commission_amount"], 39000)
        self.assertEqual(result["status"], "Eligible")

    def test_plan_d_tier_1_only(self):
        """ทดสอบ Plan D กรณียอดขายอยู่ใน Tier 1"""
        mock_data = {"sales_service_amount": 1000000} # ฐานคำนวณ = 250,000
        result = calculate_commission_data("Plan D", 0, 0, mock_data)
        # Tier 1: 250,000 * 0.7% = 1750
        self.assertAlmostEqual(result["commission_amount"], 1750)
        self.assertEqual(result["status"], "Tier 1 (0.7%)")

    def test_plan_d_tier_2(self):
        """ทดสอบ Plan D กรณียอดขายอยู่ใน Tier 2"""
        mock_data = {"sales_service_amount": 2000000} # ฐานคำนวณ = 1,250,000
        result = calculate_commission_data("Plan D", 0, 0, mock_data)
        # Tier 1 (เต็ม): 750,000 * 0.7% = 5,250
        # Tier 2 (ส่วนที่เหลือ): 500,000 * 1.0% = 5,000
        # Total = 10,250
        self.assertAlmostEqual(result["commission_amount"], 10250)
        self.assertEqual(result["status"], "Tier 2 (0.7% + 1.0%)")

# บรรทัดนี้เพื่อให้สามารถรันไฟล์นี้ได้โดยตรง
if __name__ == '__main__':
    unittest.main()
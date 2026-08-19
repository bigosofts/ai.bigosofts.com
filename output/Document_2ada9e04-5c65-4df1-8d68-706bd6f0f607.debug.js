/**
 * Reproduction script for generated DOCX output.
 * Run from the output directory:
 *   node Document_2ada9e04-5c65-4df1-8d68-706bd6f0f607.debug.js
 */
const fs = require("fs");
const path = require("path");
const { buildDocxBuffer } = require("../server");

const plan = {
  "title": "রাজেন্দ্রপুর ক্যান্টনমেন্ট পাবলিক স্কুল ও কলেজ",
  "subtitle": "০২ (দুই)টি যাত্রীবাহী লিফট ও লিফট কোর নির্মাণ, সরবরাহ, স্থাপন, পরীক্ষা, কমিশনিং ও হস্তান্তর কার্যক্রম সম্পাদন এবং নির্ধারিত সময়ের অতিরিক্ত বিলম্বের কারণে চুক্তি মোতাবেক জরিমানা আদায় সংক্রান্ত প্রতিবেদন",
  "documentType": "পরিদর্শন ও মূল্যায়ন প্রতিবেদন",
  "designPreset": "formal_administrative_register",
  "metadata": [
    "প্রকল্পের নাম: ০২টি যাত্রীবাহী লিফট ও লিফট কোর নির্মাণ ও স্থাপন",
    "সরবরাহকারী প্রতিষ্ঠান: Property Development Limited (Pvt.) Ltd. (PDL)",
    "চুক্তি মূল্য: ২,২০,০০,০০০/- টাকা (লিফট: ৯৮,০০,০০০/-, কোর: ১,২২,০০,০০০/-)",
    "কার্য শুরুর তারিখ: ১৫ জুন ২০২৫",
    "নির্ধারিত সমাপ্তির তারিখ: ১১ নভেম্বর ২০২৫",
    "বিলম্বজনিত মোট জরিমানা: ১০,৪৪,০০০/- (দশ লক্ষ চুয়াল্লিশ হাজার) টাকা"
  ],
  "sections": [
    {
      "heading": "১. প্রকল্পের সংক্ষিপ্ত বিবরণ",
      "paragraphs": [
        "রাজেন্দ্রপুর ক্যান্টনমেন্ট পাবলিক স্কুল ও কলেজের জন্য ০২ (দুই)টি SRH ব্র্যান্ডের Passenger Lift এবং সংশ্লিষ্ট Steel Structure Lift Core with Glass Cover নির্মাণ, সরবরাহ, স্থাপন, পরীক্ষা ও কমিশনিংয়ের কাজ Property Development Limited (Pvt.) Ltd. (PDL)-কে প্রদান করা হয়।",
        "কোম্পানির প্রাথমিক প্রস্তাব/Financial Offer অনুযায়ী ০২টি লিফটের মোট মূল্য ৯৮,০০,০০০/- (আটানব্বই লক্ষ) টাকা এবং Steel Structure Lift Core-এর মূল্য ১,২২,০০,০০০/- (এক কোটি বাইশ লক্ষ) টাকা নির্ধারিত ছিল। লিফটের মূল্যের মধ্যে freight, duty, installation, testing এবং site commissioning অন্তর্ভুক্ত রয়েছে।"
      ],
      "table": {
        "headers": [
          "আইটেমের বিবরণ",
          "পরিমাণ",
          "মূল্য (টাকা)"
        ],
        "rows": [
          [
            "SRH ব্র্যান্ডের Passenger Lift (Freight, Duty, Installation, Testing & Commissioning সহ)",
            "০২টি",
            "৯৮,০০,০০০/-"
          ],
          [
            "Steel Structure Lift Core with Glass Cover",
            "০১টি লট",
            "১,২২,০০,০০০/-"
          ],
          [
            "সর্বমোট চুক্তি মূল্য",
            "—",
            "২,২০,০০,০০০/-"
          ]
        ]
      }
    },
    {
      "heading": "২. প্রকল্পের সময়সীমা ও বাস্তবায়নে বিলম্ব",
      "paragraphs": [
        "চুক্তি/কার্যাদেশ অনুযায়ী প্রকল্পের কাজ ১৫ জুন ২০২৫ তারিখে শুরু হয়ে নির্ধারিত সময়ের মধ্যে, অর্থাৎ ১১ নভেম্বর ২০২৫ তারিখের মধ্যে সম্পন্ন ও হস্তান্তর করার কথা ছিল।",
        "কিন্তু নির্ধারিত সময়সীমার মধ্যে কাজ সম্পন্ন না হওয়ায় প্রকল্প বাস্তবায়নে বিলম্ব ঘটে। পরবর্তীতে কোম্পানি ১০ আগস্ট ২০২৬ তারিখে লিফটের Installation, Testing ও Commissioning সম্পন্ন হওয়ার বিষয় অবহিত করে হস্তান্তরের জন্য পত্র প্রদান করে। ফলে নির্ধারিত সময়ের তুলনায় কাজ সম্পন্ন করতে ২৭০ দিন বিলম্ব হয়েছে।"
      ]
    },
    {
      "heading": "৩. চুক্তি অনুযায়ী সরবরাহ ও কাজের অবস্থা",
      "paragraphs": [
        "চুক্তি ও Technical Specification অনুযায়ী লিফট কোম্পানির নিকট থেকে যে সকল কাজ ও সরঞ্জামাদি সরবরাহ/সম্পন্ন করার কথা ছিল, কমিটির পরিদর্শন ও যাচাইয়ে সকল নির্ধারিত কাজ ও সরঞ্জামাদি যথাযথভাবে পাওয়া গেছে। এর মধ্যে উল্লেখযোগ্য কারিগরি বৈশিষ্ট্যসমূহ নিম্নরূপ:"
      ],
      "bullets": [
        "০২টি SRH Passenger Lift (ধারণক্ষমতা: ১৬০০ কেজি / ২১ জন, ৮টি Floor/Stop, গতি: ১.০ মিটার/সেকেন্ড)",
        "Gearless Traction Machine ও 11 KW Motor",
        "VVVF Drive ও Microprocessor Controller",
        "380VAC, 3 Phase, 50Hz electrical system-এর উপযোগী Lift System",
        "ARD (Auto Rescue Device) এবং AVR (Auto Voltage Regulator)",
        "Emergency Light, Alarm ও Overload Signal",
        "3-Way Intercom ও Password System",
        "LCD Floor Indicator এবং EN-81 অনুযায়ী Safety System",
        "Steel Rope ও Speed Governor Rope",
        "নির্ধারিত specification অনুযায়ী SS ও Glass Cabin এবং Lift Door ও Landing Door",
        "Steel Structure Lift Core with Glass Cover",
        "প্রয়োজনীয় Installation, Testing ও Commissioning"
      ],
      "callout": "বিশেষ দ্রষ্টব্য: Technical Specification-এ Camera-এর ক্ষেত্রে “Only provision” উল্লেখ থাকায় CCTV Camera-কে পৃথক সরবরাহযোগ্য item হিসেবে বিবেচনা করা হয়নি। তবে Camera স্থাপনের প্রয়োজনীয় provision যথাযথভাবে রাখা হয়েছে।"
    },
    {
      "heading": "৪. ওয়ারেন্টি ও ফ্রি সার্ভিস (Warranty & Maintenance)",
      "paragraphs": [
        "চুক্তি ও Technical Specification অনুযায়ী লিফট হস্তান্তরের তারিখ হতে কোম্পানি নিম্নোক্ত সেবাসমূহ প্রদান করতে দায়বদ্ধ থাকবে:",
        "সুতরাং, লিফট হস্তান্তরের সময় উক্ত Warranty ও Free Service-এর বিষয়টি যথাযথভাবে নথিভুক্ত করে কোম্পানির নিকট হতে প্রয়োজনীয় আনুষ্ঠানিক নিশ্চয়তা গ্রহণ করা হবে।"
      ],
      "bullets": [
        "১৮ (আঠারো) মাস Warranty ও Guarantee",
        "২৪ (চব্বিশ) মাস Free Service ও Maintenance"
      ]
    },
    {
      "heading": "৫. চুক্তি মোতাবেক বিলম্বজনিত জরিমানা নির্ধারণ",
      "paragraphs": [
        "যদিও কোম্পানি কর্তৃক চুক্তি ও Technical Specification অনুযায়ী নির্ধারিত সকল কাজ ও সরঞ্জামাদি সরবরাহ এবং Installation, Testing ও Commissioning সম্পন্ন করা হয়েছে, তথাপি নির্ধারিত সময়সীমার মধ্যে প্রকল্পের কাজ সম্পন্ন করা হয়নি। চুক্তির শর্ত মোতাবেক বিলম্বজনিত জরিমানা কার্যকর হবে।",
        "নির্ধারিত শেষ তারিখ ১২ নভেম্বর ২০২৫ এবং হস্তান্তরের তারিখ ৩১ জুলাই ২০২৬ বিবেচনায় মোট ২৬১ দিনের বিলম্ব হিসাব করা হয়েছে।"
      ],
      "table": {
        "headers": [
          "বিবরণ",
          "দৈনিক জরিমানার হার",
          "বিলম্বিত কার্যদিবস",
          "মোট জরিমানার পরিমাণ (টাকা)"
        ],
        "rows": [
          [
            "Lift Core বিলম্ব জরিমানা",
            "২,০০০/- টাকা",
            "২৬১ দিন",
            "৫,২২,০০০/-"
          ],
          [
            "০২টি Lift বিলম্ব জরিমানা",
            "২,০০০/- টাকা",
            "২৬১ দিন",
            "৫,২২,০০০/-"
          ],
          [
            "সর্বমোট প্রযোজ্য জরিমানা",
            "৪,০০০/- টাকা",
            "২৬১ দিন",
            "১০,৪৪,০০০/-"
          ]
        ]
      },
      "callout": "জরিমানা হিসাব সূত্র: ২৬১ দিন × ৪,০০০/- টাকা = ১০,৪৪,০০০/- (দশ লক্ষ চুয়াল্লিশ হাজার) টাকা মাত্র কোম্পানির চূড়ান্ত বিল হতে কর্তনযোগ্য।"
    },
    {
      "heading": "৬. কমিটির মতামত ও সুপারিশ",
      "paragraphs": [
        "সার্বিক পর্যালোচনায় প্রতীয়মান হয় যে, লিফট কোম্পানি কর্তৃক চুক্তি/Technical Specification অনুযায়ী নির্ধারিত সকল কাজ ও সরঞ্জামাদি সরবরাহ এবং স্থাপন করা হয়েছে এবং লিফটসমূহ Testing ও Commissioning সম্পন্ন করে ব্যবহারের উপযোগী করা হয়েছে।",
        "অতএব, লিফটের কাজ ও সরবরাহকৃত সরঞ্জামাদি বিষয়ে কমিটির কোনো আপত্তি নেই এবং প্রযোজ্য জরিমানা বাবদ ১০,৪৪,০০০/- (দশ লক্ষ চুয়াল্লিশ হাজার) টাকা কর্তন সাপেক্ষে লিফটসমূহ চূড়ান্তভাবে হস্তান্তর গ্রহণের জন্য উপযুক্ত মর্মে কমিটি সর্বসম্মত মতামত ও সুপারিশ প্রদান করছে।"
      ]
    },
    {
      "heading": "৭. পরিদর্শন ও হস্তান্তর মূল্যায়ন কমিটির স্বাক্ষর",
      "signatures": [
        {
          "role": "সভাপতি",
          "name": "সভাপতি, পরিদর্শন ও মূল্যায়ন কমিটি"
        },
        {
          "role": "সদস্য",
          "name": "সদস্য ১"
        },
        {
          "role": "সদস্য",
          "name": "সদস্য ২"
        },
        {
          "role": "সদস্য",
          "name": "সদস্য ৩"
        },
        {
          "role": "সদস্য",
          "name": "সদস্য ৪"
        }
      ]
    }
  ],
  "imagePlacements": []
};
const images = [];

(async function reproduce() {
  try {
    const buffer = await buildDocxBuffer(plan, images);
    const outPath = path.join(__dirname, "Document_2ada9e04-5c65-4df1-8d68-706bd6f0f607.repro.docx");
    fs.writeFileSync(outPath, buffer);
    console.log("Reproduced DOCX saved to:", outPath);
  } catch (error) {
    console.error("Reproduction failed:", error);
    process.exit(1);
  }
})();

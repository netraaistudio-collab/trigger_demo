
import { schedules, wait } from "@trigger.dev/sdk";
import ExcelJS from "exceljs";

// ─── Types ────────────────────────────────────────────────────────────────────

interface PlacesTextSearchResult {
  place_id: string;
  name: string;
}

interface PlacesTextSearchResponse {
  results: PlacesTextSearchResult[];
  next_page_token?: string;
  status: string;
}

interface PlaceDetailsResponse {
  result: {
    name?: string;
    formatted_address?: string;
    formatted_phone_number?: string;
    website?: string;
    rating?: number;
    user_ratings_total?: number;
    url?: string;
  };
  status: string;
}

interface DentalLead {
  name: string;
  address: string;
  phone: string;
  website: string;
  rating: string;
  reviews: string;
  mapsUrl: string;
}

// ─── Scheduled Task ───────────────────────────────────────────────────────────

export const mumbaiDentalLeads = schedules.task({
  id: "mumbai-dental-leads",
  cron: "30 3 * * 0", // Every Sunday at 9:00am IST (3:30am UTC)

  run: async () => {
    // Validate env vars
    const googleApiKey = process.env.GOOGLE_PLACES_API_KEY;
    if (!googleApiKey) throw new Error("GOOGLE_PLACES_API_KEY is not set");

    const resendApiKey = process.env.RESEND_API_KEY;
    if (!resendApiKey) throw new Error("RESEND_API_KEY is not set");

    const fromEmail = process.env.RESEND_FROM_EMAIL;
    if (!fromEmail) throw new Error("RESEND_FROM_EMAIL is not set");

    const toEmail = process.env.LEAD_EMAIL_TO;
    if (!toEmail) throw new Error("LEAD_EMAIL_TO is not set");

    // ── Step 1: Fetch page 1 of dental clinics in Mumbai (20 results) ──────────
    console.log("Fetching dental clinics in Mumbai — page 1...");

    const page1Url =
      `https://maps.googleapis.com/maps/api/place/textsearch/json` +
      `?query=dental+clinic+Mumbai&key=${googleApiKey}`;

    const page1Res = await fetch(page1Url);
    const page1Data = (await page1Res.json()) as PlacesTextSearchResponse;

    if (page1Data.status !== "OK") {
      throw new Error(`Places Text Search failed: ${page1Data.status}`);
    }

    const placeIds: string[] = page1Data.results.map((r) => r.place_id);
    console.log(`Page 1: got ${placeIds.length} results`);

    // ── Step 2: Fetch page 2 for 5 more results ────────────────────────────────
    if (page1Data.next_page_token && placeIds.length < 25) {
      // Must wait before using next_page_token or the API returns INVALID_REQUEST
      await wait.for({ seconds: 2 });

      console.log("Fetching page 2...");
      const page2Url =
        `https://maps.googleapis.com/maps/api/place/textsearch/json` +
        `?pagetoken=${page1Data.next_page_token}&key=${googleApiKey}`;

      const page2Res = await fetch(page2Url);
      const page2Data = (await page2Res.json()) as PlacesTextSearchResponse;

      if (page2Data.status === "OK") {
        const needed = 25 - placeIds.length;
        const extra = page2Data.results.slice(0, needed).map((r) => r.place_id);
        placeIds.push(...extra);
        console.log(`Page 2: got ${extra.length} more — total ${placeIds.length}`);
      }
    }

    // Cap at 25
    const targetIds = placeIds.slice(0, 25);

    // ── Step 3: Fetch Place Details for each clinic ────────────────────────────
    console.log(`Fetching details for ${targetIds.length} clinics...`);

    const leads: DentalLead[] = [];

    for (const placeId of targetIds) {
      const detailsUrl =
        `https://maps.googleapis.com/maps/api/place/details/json` +
        `?place_id=${placeId}` +
        `&fields=name,formatted_address,formatted_phone_number,website,rating,user_ratings_total,url` +
        `&key=${googleApiKey}`;

      const detailsRes = await fetch(detailsUrl);
      const detailsData = (await detailsRes.json()) as PlaceDetailsResponse;

      if (detailsData.status !== "OK") {
        console.warn(`Details fetch failed for ${placeId}: ${detailsData.status}`);
        continue;
      }

      const r = detailsData.result;
      leads.push({
        name: r.name ?? "Unknown",
        address: r.formatted_address ?? "—",
        phone: r.formatted_phone_number ?? "—",
        website: r.website ?? "No website",
        rating: r.rating != null ? String(r.rating) : "—",
        reviews: r.user_ratings_total != null ? String(r.user_ratings_total) : "—",
        mapsUrl: r.url ?? "—",
      });
    }

    console.log(`Collected ${leads.length} leads`);

    // ── Step 4: Build Excel file ───────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Mumbai Dental Leads");

    // Header row
    sheet.columns = [
      { header: "Clinic Name", key: "name", width: 35 },
      { header: "Address", key: "address", width: 50 },
      { header: "Phone", key: "phone", width: 20 },
      { header: "Website", key: "website", width: 40 },
      { header: "Rating", key: "rating", width: 10 },
      { header: "Reviews", key: "reviews", width: 12 },
      { header: "Google Maps URL", key: "mapsUrl", width: 55 },
    ];

    // Style header row
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1A73E8" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 20;

    // Add data rows
    for (const lead of leads) {
      sheet.addRow(lead);
    }

    // Freeze header row
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    // Convert to base64 buffer for email attachment
    const buffer = await workbook.xlsx.writeBuffer();
    const base64 = Buffer.from(buffer).toString("base64");

    // ── Step 5: Send email via Resend ─────────────────────────────────────────
    const today = new Date().toLocaleDateString("en-IN", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      timeZone: "Asia/Kolkata",
    });

    const noWebsiteCount = leads.filter((l) => l.website === "No website").length;

    const emailRes = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${resendApiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: fromEmail,
        to: toEmail,
        subject: `Mumbai Dental Leads — ${today}`,
        html: `
          <h2>Mumbai Dental Leads — ${today}</h2>
          <p>Your weekly batch of <strong>${leads.length} dental clinic leads</strong> from Mumbai is attached.</p>
          <ul>
            <li><strong>${noWebsiteCount}</strong> clinics have <strong>no website</strong> — highest priority leads</li>
            <li><strong>${leads.length - noWebsiteCount}</strong> clinics have an existing website</li>
          </ul>
          <p>Open the attached Excel file to view all leads with contact details.</p>
          <hr/>
          <p style="color:#888;font-size:12px;">This email was sent automatically every Sunday by your Trigger.dev automation.</p>
        `,
        attachments: [
          {
            filename: `mumbai-dental-leads-${new Date().toISOString().slice(0, 10)}.xlsx`,
            content: base64,
          },
        ],
      }),
    });

    if (!emailRes.ok) {
      const err = await emailRes.text();
      throw new Error(`Resend API error: ${emailRes.status} — ${err}`);
    }

    console.log(`Email sent to ${toEmail} with ${leads.length} leads`);

    return {
      leadsFound: leads.length,
      noWebsiteCount,
      emailSent: true,
      sentTo: toEmail,
    };
  },
});

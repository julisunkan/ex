export default function PrivacyPage() {
  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col">
      <header className="flex items-center gap-3 px-4 py-3 bg-white border-b border-border shadow-sm shrink-0">
        <a
          href="/"
          className="flex items-center gap-1 text-sm font-semibold text-muted-foreground hover:text-foreground px-2.5 py-1.5 rounded-lg hover:bg-muted transition-colors"
        >
          <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
            <polyline points="15 18 9 12 15 6" />
          </svg>
          Back
        </a>
        <h1 className="text-sm font-bold text-foreground">Privacy Policy</h1>
      </header>

      <div className="flex-1 overflow-y-auto px-5 py-6 space-y-5 text-sm leading-relaxed text-foreground">
        <p className="text-xs text-muted-foreground">Last updated: June 2026</p>

        <p>
          This Privacy Policy describes how <strong>Bank Statement Analyzer</strong> ("the Add-in", "we", "us") collects, uses, and handles your information when you use our Microsoft Excel Add-in service.
        </p>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">1. Information We Collect</h2>
          <p><strong>Information you provide directly:</strong></p>
          <ul className="list-disc pl-5 space-y-1">
            <li><strong>Email address</strong> — optionally provided during payment to receive your license key and renewal reminders.</li>
            <li><strong>License key</strong> — generated when you complete a payment, stored to validate your subscription.</li>
            <li><strong>Support ticket data</strong> — your name, email, and message content when you submit a support request.</li>
            <li><strong>Blockchain transaction hash</strong> — submitted by you to verify your payment on-chain.</li>
          </ul>
          <p className="mt-2"><strong>Information we do NOT collect:</strong></p>
          <ul className="list-disc pl-5 space-y-1">
            <li>Your bank statement data, transaction records, or financial information — all spreadsheet data is processed entirely within your local Excel session and is never transmitted to our servers.</li>
            <li>Passwords, banking credentials, or account numbers.</li>
            <li>Device identifiers, IP addresses, or browser fingerprints.</li>
          </ul>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">2. How We Use Your Information</h2>
          <p>We use the information we collect only to:</p>
          <ul className="list-disc pl-5 space-y-1">
            <li>Verify your payment and issue your license key;</li>
            <li>Validate your subscription status when you use the Add-in;</li>
            <li>Send subscription renewal reminder emails (only if you provided your email and reminders are enabled);</li>
            <li>Respond to support tickets you submit;</li>
            <li>Detect and prevent fraud or abuse of the licensing system.</li>
          </ul>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">3. Data Storage &amp; Security</h2>
          <p>
            License and ticket data is stored securely on our server. We implement reasonable technical measures to protect your information. Your financial data (bank statements) is <strong>never</strong> stored on our servers — it remains in your Excel workbook at all times.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">4. Data Sharing</h2>
          <p>
            We do not sell, trade, rent, or share your personal information with third parties, except:
          </p>
          <ul className="list-disc pl-5 space-y-1">
            <li>As required by applicable law or legal process;</li>
            <li>To protect the rights, property, or safety of Bank Analyzer Pro or our users;</li>
            <li>With blockchain networks (Tron, BSC, Ethereum) solely to verify payment transactions — only the transaction hash you provide is used.</li>
          </ul>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">5. Cookies &amp; Tracking</h2>
          <p>
            Bank Statement Analyzer does not use cookies, analytics trackers, or any third-party tracking technologies within the Add-in. Your usage of the Add-in is not monitored or recorded.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">6. Data Retention</h2>
          <p>
            License records are retained for as long as necessary to validate subscriptions and prevent fraud. Support ticket records are retained for up to 12 months. You may request deletion of your data by contacting us via the Support page.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">7. Your Rights</h2>
          <p>
            Depending on your jurisdiction, you may have the right to access, correct, or delete personal information we hold about you. To exercise these rights, please submit a support ticket via the Add-in's Support page.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">8. Children's Privacy</h2>
          <p>
            Bank Statement Analyzer is not directed at children under the age of 13. We do not knowingly collect personal information from children.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">9. Changes to This Policy</h2>
          <p>
            We may update this Privacy Policy from time to time. The updated version will be available within the Add-in with a revised "Last updated" date. Continued use of Bank Statement Analyzer after changes constitutes acceptance of the updated policy.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">10. Contact</h2>
          <p>
            For privacy-related questions or requests, please submit a support ticket via the Add-in's Support page.
          </p>
        </section>
      </div>
    </div>
  );
}

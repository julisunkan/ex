export default function EulaPage() {
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
        <h1 className="text-sm font-bold text-foreground">End User License Agreement</h1>
      </header>

      <div className="flex-1 overflow-y-auto px-5 py-6 space-y-5 text-sm leading-relaxed text-foreground">
        <p className="text-xs text-muted-foreground">Last updated: June 2026</p>

        <p>
          This End User License Agreement ("Agreement") is a legal agreement between you ("User") and Bank Analyzer Pro ("Licensor") for the use of <strong>Bank Statement Analyzer</strong> ("the Add-in"), a Microsoft Excel Add-in.
        </p>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">1. License Grant</h2>
          <p>
            Subject to the terms of this Agreement, the Licensor grants you a limited, non-exclusive, non-transferable, revocable license to install and use the Add-in solely for your personal or internal business purposes.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">2. Restrictions</h2>
          <p>You may not:</p>
          <ul className="list-disc pl-5 space-y-1 text-sm">
            <li>Sublicense, sell, resell, transfer, assign, or otherwise commercially exploit the Add-in;</li>
            <li>Reverse engineer, decompile, disassemble, or attempt to derive the source code of the Add-in;</li>
            <li>Share, distribute, or make your license key available to any third party;</li>
            <li>Use the Add-in for any unlawful purpose or in violation of any applicable laws;</li>
            <li>Remove or alter any proprietary notices, labels, or marks on the Add-in.</li>
          </ul>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">3. Subscription &amp; Payment</h2>
          <p>
            Pro features require a paid subscription. Payments are processed in USDT cryptocurrency via the blockchain network you select. All sales are final — no refunds are issued once a license key has been issued and verified. License keys are non-transferable and tied to the purchasing user.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">4. Intellectual Property</h2>
          <p>
            The Add-in, including all content, features, and functionality, is and remains the exclusive property of the Licensor. This Agreement does not convey any ownership rights in the Add-in to you.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">5. Disclaimer of Warranties</h2>
          <p>
            THE ADD-IN IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED. THE LICENSOR DOES NOT WARRANT THAT THE ADD-IN WILL BE UNINTERRUPTED, ERROR-FREE, OR FREE OF VIRUSES. YOUR USE OF THE ADD-IN IS AT YOUR SOLE RISK.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">6. Limitation of Liability</h2>
          <p>
            TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, THE LICENSOR SHALL NOT BE LIABLE FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES, INCLUDING LOSS OF PROFITS OR DATA, ARISING OUT OF OR IN CONNECTION WITH YOUR USE OF THE ADD-IN.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">7. Termination</h2>
          <p>
            This Agreement is effective until terminated. Your rights under this Agreement will terminate automatically if you fail to comply with any of its terms. Upon termination, you must cease all use of the Add-in and delete all copies in your possession.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">8. Governing Law</h2>
          <p>
            This Agreement shall be governed by and construed in accordance with applicable international laws. Any disputes arising under this Agreement shall be resolved through binding arbitration.
          </p>
        </section>

        <section className="space-y-1.5">
          <h2 className="font-bold text-base">9. Contact</h2>
          <p>
            For questions about this Agreement, please submit a support ticket via the Add-in's Support page.
          </p>
        </section>

        <p className="text-xs text-muted-foreground pt-2 border-t border-border">
          By installing or using Bank Statement Analyzer, you agree to be bound by the terms of this Agreement.
        </p>
      </div>
    </div>
  );
}

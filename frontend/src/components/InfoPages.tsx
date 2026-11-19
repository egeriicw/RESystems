export function AboutPage() {
  return (
    <article className="card prose">
      <h2>About</h2>
      <p>This is a web port of the <strong>Renewable Energy Systems</strong> Visual Basic 6 program (see
        <code> Resources/</code> in the repository). It simulates a year of hourly solar collector and storage
        performance from typical-year weather data and reports monthly averages, collector efficiency and solar fraction.</p>
      <h3>Models</h3>
      <ul>
        <li><strong>Collector useful energy</strong> – Hottel–Whillier–Bliss gain at a fixed inlet temperature.</li>
        <li><strong>House heating</strong> – collector charges a tank directly; the tank meets the house load when warmer than the room.</li>
        <li><strong>Domestic hot water</strong> – collector heat exchanger to a tank serving an hourly hot-water draw profile.</li>
        <li><strong>House heating with fan coil</strong> – tank serves the house through a fan coil above a minimum supply temperature.</li>
        <li><strong>Fan coil + DHW</strong> – combined space and water heating from one tank.</li>
      </ul>
      <h3>Method</h3>
      <p>Solar position and incidence angle follow Duffie &amp; Beckman. Global horizontal radiation is split into beam and
        diffuse with the Erbs correlation, and radiation on the tilted collector uses the HDKR anisotropic sky model.
        A glazing-dependent incidence angle modifier is applied. Storage is a single fully mixed tank stepped hourly.</p>
    </article>
  );
}

export function HelpPage() {
  return (
    <article className="card prose">
      <h2>Help</h2>
      <ol>
        <li><strong>System diagram</strong> – click a component in the library (or drag it onto the canvas), then wire it up by dragging from a coloured port to a port of the same colour. Once everything is connected the system is simulated. Click any component or wire to edit its parameters; “Save &amp; rerun” updates the results. Start from a template to see the original program’s five systems.</li>
        <li><strong>Weather</strong> – a synthetic Dayton, OH year is loaded by default. Upload a TMY2 file or generate a synthetic year for another site.</li>
        <li><strong>Simulator → Inputs</strong> – pick a model type, then set collector and system parameters. Only the parameters the chosen model uses are shown.</li>
        <li><strong>Sites</strong> – add a building by clicking the map and filling in its features. The nearest weather dataset is assigned automatically.</li>
        <li><strong>Scenarios</strong> – name a set of inputs in the side panel and choose “Save &amp; run” to store it and its results. Tick two scenarios in the Scenarios view to compare them.</li>
        <li><strong>Run simulation</strong> (nothing is stored), then review the <strong>Monthly table</strong>, <strong>Trends</strong> (monthly, daily or hourly) and the hour-by-day <strong>Heat map</strong>.</li>
      </ol>
      <h3>Units</h3>
      <p>Energy is in kJ (hourly values in kJ/hr; daily and monthly values are average daily totals in kJ/day). Radiation is per m² of collector. Temperatures are °C.
        1 kWh = 3,600 kJ.</p>
    </article>
  );
}

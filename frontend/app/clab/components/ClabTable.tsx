import "./ClabTable.css";

const rows = [
  {
    label: "Mon",
    volume: 118,
    resolved: 82,
    risk: 36,
  },
  {
    label: "Tue",
    volume: 164,
    resolved: 118,
    risk: 46,
  },
  {
    label: "Wed",
    volume: 142,
    resolved: 109,
    risk: 33,
  },
  {
    label: "Thu",
    volume: 201,
    resolved: 161,
    risk: 40,
  },
  {
    label: "Fri",
    volume: 188,
    resolved: 154,
    risk: 34,
  },
];

export default function ClabTable() {
  return (
    <section className="clab-table-card">
      <div className="clab-table-header">
        <div>
          <h3>Signal distribution</h3>

          <p>
            Edits in the table update every chart instantly.
          </p>
        </div>

        <span className="live-update">
          Live update
        </span>
      </div>

      <div className="table-wrapper">
        <table className="clab-table">
          <thead>
            <tr>
              <th>LABEL</th>
              <th>VOLUME</th>
              <th>RESOLVED</th>
              <th>RISK</th>
            </tr>
          </thead>

          <tbody>
            {rows.map((row) => (
              <tr key={row.label}>
                <td>{row.label}</td>
                <td>{row.volume}</td>
                <td>{row.resolved}</td>
                <td>{row.risk}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
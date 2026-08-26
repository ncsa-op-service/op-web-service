import "./BarChart.css";

const data = [
  { label: "Mon", volume: 118, resolved: 82 },
  { label: "Tue", volume: 164, resolved: 118 },
  { label: "Wed", volume: 142, resolved: 109 },
  { label: "Thu", volume: 201, resolved: 161 },
  { label: "Fri", volume: 188, resolved: 154 },
];

export default function BarChart() {
  return (
    <div className="clab-bar-chart">
      <div className="bar-y-axis">
        <span>220</span>
        <span>165</span>
        <span>110</span>
        <span>55</span>
        <span>0</span>
      </div>

      <div className="bar-grid">
        <div />
        <div />
        <div />
        <div />
        <div />
      </div>

      <div className="bar-data-area">
        {data.map((item) => (
          <div className="bar-group" key={item.label}>
            <div className="bar-columns">
              <div
                className="bar-column volume"
                style={{
                  height: `${(item.volume / 220) * 250}px`,
                }}
              />

              <div
                className="bar-column resolved"
                style={{
                  height: `${(item.resolved / 220) * 250}px`,
                }}
              />
            </div>

            <span className="bar-label">{item.label}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
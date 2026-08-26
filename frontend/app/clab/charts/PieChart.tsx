import "./PieChart.css";

export default function PieChart() {
  return (
    <div className="clab-pie-container">
      <div className="clab-pie-chart">
        <div className="clab-pie-hole" />
      </div>

      <span className="clab-pie-label pie-tue">
        Tue:164
      </span>

      <span className="clab-pie-label pie-mon">
        Mon:118
      </span>

      <span className="clab-pie-label pie-fri">
        Fri:188
      </span>

      <span className="clab-pie-label pie-thu">
        Thu:201
      </span>

      <span className="clab-pie-label pie-wed">
        Wed:142
      </span>
    </div>
  );
}
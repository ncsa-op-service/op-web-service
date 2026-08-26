import "./BarChart.css";

export default function BarChart() {
  return (
    <div className="bar-chart-wrapper">
      <div className="bar-chart-area">
        <div className="bar-item">
          <div className="bar bar-one"></div>
          <span>Tool 1</span>
        </div>

        <div className="bar-item">
          <div className="bar bar-two"></div>
          <span>Tool 2</span>
        </div>

        <div className="bar-item">
          <div className="bar bar-three"></div>
          <span>Tool 3</span>
        </div>

        <div className="bar-item">
          <div className="bar bar-four"></div>
          <span>Tool 4</span>
        </div>
      </div>
    </div>
  );
}
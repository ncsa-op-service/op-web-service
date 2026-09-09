import "./BarChart.css";

type ChartItem = {
  label: string;
  value: number;
};

type BarChartProps = {
  data: ChartItem[];
};

export default function BarChart({
  data,
}: BarChartProps) {
  const maxValue = Math.max(
    ...data.map((item) => item.value),
    1
  );

  return (
    <div className="bar-chart-wrapper">
      <div className="bar-chart-area">
        {data.map((item) => {
          const height =
            (item.value / maxValue) * 100;

          return (
            <div
              className="bar-item"
              key={item.label}
            >
              <div className="bar-value">
                {item.value}
              </div>

              <div className="bar-track">
                <div
                  className="bar"
                  style={{
                    height: `${height}%`,
                  }}
                />
              </div>

              <span>{item.label}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
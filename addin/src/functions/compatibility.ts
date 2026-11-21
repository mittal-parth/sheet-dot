function echo(value: string): string {
  return value;
}

function stream(value: string, invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setTimeout(() => invocation.setResult(value), 1500);
  invocation.onCanceled = () => clearTimeout(timer);
  invocation.setResult("Working…");
}

function spill(): string[][] {
  return [["Row 1", "Spill worked"], ["Row 2", "2 columns"]];
}

CustomFunctions.associate("ECHO", echo);
CustomFunctions.associate("STREAM", stream);
CustomFunctions.associate("SPILL", spill);

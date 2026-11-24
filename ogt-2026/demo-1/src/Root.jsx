import {Composition} from 'remotion';
import {Demo1, DURATION, FPS, HEIGHT, WIDTH} from './Demo1';

export const Root = () => {
  return (
    <Composition
      id="Demo1"
      component={Demo1}
      durationInFrames={DURATION}
      fps={FPS}
      width={WIDTH}
      height={HEIGHT}
    />
  );
};
